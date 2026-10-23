using System.Reflection;
using System.Text.Json;
using BcWordLayout.McpHost;
using BcWordLayout.McpHost.Tools;
using BcWordLayout.Render;
using ModelContextProtocol.Server;

namespace BcWordLayout.Tests;

/// <summary>
/// Pins what the call boundary returns for calls that never reach a tool body — the cases README's "Response
/// envelope" section and ADR-0004 describe — by driving the real server pipeline (<see cref="McpHostPipeline"/>:
/// SDK framing, tool lookup, argument binding, and <c>ToolCallBoundary</c>'s filter). Every one of these used
/// to be undocumented or mis-coded (GitHub issue #97): a wrong-typed argument and an unknown tool came back as
/// <c>internal_error</c>, and an explicit null for an enum-like string crashed the tool body. These tests are
/// also what re-verifies the boundary on an SDK upgrade: they fail if a new <c>ModelContextProtocol</c>
/// version changes how binding failures reach the filter.
/// </summary>
/// <remarks>Joins the preview-converter-seam collection because the null sweep calls <c>preview_layout</c>
/// (against a missing layout, so it never converts — and with the converter seam swapped for a fake that
/// fails the test if it is ever asked to; see <see cref="PreviewConverterSeamCollection"/>).</remarks>
[Collection("preview-converter-seam")]
public sealed class McpHostPipelineTests(McpHostPipeline pipeline) : IClassFixture<McpHostPipeline>, IDisposable
{
    private readonly string workDir = Directory.CreateDirectory(
        Path.Combine(Path.GetTempPath(), $"bcwl-pipeline-{Guid.NewGuid():N}")).FullName;

    public void Dispose()
    {
        try
        {
            Directory.Delete(workDir, recursive: true);
        }
        catch (IOException)
        {
            // Best-effort temp cleanup.
        }
    }

    private string CopyOfCorpus(string corpusFile)
    {
        var path = Path.Combine(workDir, $"{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(corpusFile), path);
        return path;
    }

    /// <summary>An absolute path that does not exist: enough for calls that must fail before touching a file.</summary>
    private string MissingFile(string extension = ".docx") => Path.Combine(workDir, $"missing-{Guid.NewGuid():N}{extension}");

    private static string J(string value) => McpHostPipeline.Json(value);

    // ---- a well-formed call passes through untouched ----

    [Fact]
    public async Task A_well_formed_call_passes_through_the_filter_unchanged()
    {
        var result = await pipeline.CallAsync("get_layout_info", $$"""{"layoutPath":{{J(CopyOfCorpus(Corpus.SalesInvoice))}}}""");

        Assert.NotEqual(true, result.IsError);
        Assert.True(result.Ok, result.RawText);
        Assert.Equal("1306", result.Data.GetProperty("report").GetProperty("reportId").GetString());
    }

    // ---- server identity (GitHub issue #112) ----

    [Fact]
    public void The_handshake_reports_the_server_name_and_its_product_version_not_the_assembly_version()
    {
        var assembly = typeof(ServerIdentity).Assembly;
        var informational = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()!.InformationalVersion;

        var serverInfo = pipeline.Client.ServerInfo;

        Assert.Equal("BcWordLayout.McpHost", serverInfo.Name);
        Assert.Equal(ServerIdentity.WithoutBuildMetadata(informational), serverInfo.Version);
        // MAJOR.MINOR.PATCH plus an optional pre-release tag (1.1.0, 1.2.0-rc.1) - never the numeric four-part
        // assembly version (1.1.0.0) that an RC shares with its release, and never the +commit build metadata.
        Assert.Matches(@"^\d+\.\d+\.\d+(-[0-9A-Za-z.-]+)?$", serverInfo.Version);
        Assert.NotEqual(assembly.GetName().Version!.ToString(), serverInfo.Version);
    }

    [Theory]
    [InlineData("1.1.0", "1.1.0")]
    [InlineData("1.2.0-rc.1", "1.2.0-rc.1")]
    [InlineData("1.2.0-rc.1+3f2a9c1d", "1.2.0-rc.1")]
    [InlineData("1.1.0+3f2a9c1d.dirty", "1.1.0")]
    public void WithoutBuildMetadata_strips_only_the_plus_suffix(string informational, string expected) =>
        Assert.Equal(expected, ServerIdentity.WithoutBuildMetadata(informational));

    // ---- missing required argument ----

    [Fact]
    public async Task A_missing_required_argument_is_invalid_argument_naming_it_and_the_required_set()
    {
        var result = await pipeline.CallAsync(
            "insert_field", $$"""{"layoutPath":{{J(MissingFile())}},"locationType":"documentEnd"}""");

        Assert.True(result.IsError);
        result.AssertFailure("invalid_argument", "insert_field", "'field'");
        Assert.Contains("layoutPath, field, locationType", result.Hint, StringComparison.Ordinal);
    }

    // ---- wrong JSON type ----

    [Theory]
    [InlineData("insert_field", """ "field":"/Header/CustomerAddress1","locationType":"tableCell","tableIndex":"abc" """, "tableIndex", "an integer or null", "the string \"abc\"")]
    [InlineData("remove_control", """ "controlId":1,"keepText":"true" """, "keepText", "a boolean (true or false)", "the string \"true\"")]
    [InlineData("insert_table", """ "rows":1.5,"columns":2,"locationType":"documentEnd" """, "rows", "an integer", "the number 1.5")]
    [InlineData("set_cell_format", """ "tableIndex":0,"row":0,"edges":"top","borderSize":"thick" """, "borderSize", "an integer or null", "the string \"thick\"")]
    [InlineData("remove_control", """ "controlId":[1] """, "controlId", "an integer", "an array")]
    public async Task A_wrong_typed_argument_is_invalid_argument_naming_the_parameter_and_its_type(
        string tool, string otherArguments, string parameter, string expected, string sent)
    {
        var result = await pipeline.CallAsync(
            tool, $$"""{"layoutPath":{{J(MissingFile())}},{{otherArguments}}}""");

        Assert.True(result.IsError);
        result.AssertFailure("invalid_argument", $"Argument '{parameter}' of {tool}", $"must be {expected}", $"sent {sent}");
        Assert.Contains($"'{parameter}'", result.Hint, StringComparison.Ordinal);
    }

    // ---- explicit null for an enum-like string ----

    [Fact]
    public async Task An_explicit_null_layoutPart_is_the_documented_default_body()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);

        var result = await pipeline.CallAsync(
            "insert_text", $$"""{"layoutPath":{{J(path)}},"text":"x","locationType":"documentEnd","layoutPart":null}""");

        Assert.True(result.Ok, result.RawText);
        Assert.Equal("document.xml", result.Data.GetProperty("part").GetString());
    }

    [Theory]
    [InlineData("insert_repeater_table", """ "dataItem":"/Header/Line","columns":"ItemNo_Line","locationType":"documentEnd","borders":null """)]
    [InlineData("insert_table", """ "rows":1,"columns":2,"locationType":"documentEnd","borders":null """)]
    [InlineData("set_cell_format", """ "tableIndex":0,"row":0,"edges":"top","borderStyle":null """)]
    [InlineData("validate_layout", """ "level":null """)]
    public async Task An_explicit_null_for_an_optional_enum_like_string_is_its_documented_default(
        string tool, string otherArguments)
    {
        var result = await pipeline.CallAsync(
            tool, $$"""{"layoutPath":{{J(CopyOfCorpus(Corpus.SalesInvoice))}},{{otherArguments}}}""");

        Assert.True(result.Ok, result.RawText);
    }

    [Theory]
    [InlineData("insert_field", """ "field":"/Header/CustomerAddress1","locationType":null """, "locationType", "locationType must be one of")]
    [InlineData("insert_column", """ "tableIndex":0,"mode":null """, "mode", "insert_column's mode must be")]
    public async Task An_explicit_null_for_a_required_enum_like_string_is_invalid_argument_naming_it(
        string tool, string otherArguments, string parameter, string hintFragment)
    {
        var result = await pipeline.CallAsync(
            tool, $$"""{"layoutPath":{{J(CopyOfCorpus(Corpus.SalesInvoice))}},{{otherArguments}}}""");

        result.AssertFailure("invalid_argument", $"{parameter} is required (got null)");
        Assert.Contains(hintFragment, result.Hint, StringComparison.Ordinal);
    }

    [Fact]
    public void An_explicit_null_converter_is_the_documented_default_auto()
    {
        // Pinned on the parser rather than through preview_layout: a passing converter argument leads straight
        // into a real PDF conversion, which this suite never starts.
        Assert.True(LifecycleTools.TryParseConverterKind(null, out var kind));
        Assert.Equal(PdfConverterKind.Auto, kind);
    }

    /// <summary>
    /// Sends an explicit JSON null for EVERY string parameter of EVERY tool, one at a time, with the other
    /// required arguments set to plausible values: none may come back as <c>internal_error</c> (the "bug in the
    /// tool" code). New tools and parameters are covered automatically; a new REQUIRED parameter name needs a
    /// baseline value below, and the test says so. <c>preview_layout</c> is always pointed at a missing layout
    /// and the converter seam at a fake that records any use, so the sweep can never start a conversion.
    /// </summary>
    [Fact]
    public async Task No_string_parameter_of_any_tool_turns_an_explicit_null_into_internal_error()
    {
        var converterUses = 0;
        LifecycleTools.SelectConverter = _ =>
        {
            converterUses++;
            return new FakePdfConverter();
        };

        var failures = new List<string>();
        try
        {
            foreach (var (tool, method) in ToolMethods())
            {
                foreach (var nulled in method.GetParameters().Where(p => p.ParameterType == typeof(string)))
                {
                    var arguments = new Dictionary<string, object?> { [nulled.Name!] = null };
                    foreach (var p in method.GetParameters().Where(p => !p.HasDefaultValue && p.Name != nulled.Name))
                    {
                        arguments[p.Name!] = BaselineValue(tool, p);
                    }

                    var result = await pipeline.CallAsync(tool, JsonSerializer.Serialize(arguments));
                    if (result.Code == "internal_error")
                    {
                        failures.Add($"{tool}.{nulled.Name}=null -> internal_error: {result.Message}");
                    }
                }
            }
        }
        finally
        {
            LifecycleTools.SelectConverter = PdfConverterFactory.Select;
        }

        Assert.Equal(0, converterUses);
        Assert.True(failures.Count == 0, string.Join(Environment.NewLine, failures));
    }

    private object BaselineValue(string tool, ParameterInfo parameter)
    {
        if (parameter.ParameterType == typeof(int))
        {
            return parameter.Name switch
            {
                "rows" or "columns" or "toColumn" => 1,
                _ => 0,
            };
        }

        Assert.True(parameter.ParameterType == typeof(string), $"{tool}.{parameter.Name}: no baseline for type {parameter.ParameterType}.");
        return parameter.Name switch
        {
            "layoutPath" when tool == "preview_layout" => MissingFile(),
            "layoutPath" or "source" or "schemaSource" or "newSchemaSource" => CopyOfCorpus(Corpus.SalesInvoice),
            "outputPath" => Path.Combine(workDir, $"out-{Guid.NewGuid():N}.docx"),
            "pdfPath" => MissingFile(".pdf"),
            "field" => "/Header/CustomerAddress1",
            "label" => "/Header/Contact_Lbl",
            "picture" => "/Header/CompanyPicture",
            "dataItem" => "/Header/Line",
            "columns" => "ItemNo_Line",
            "locationType" => "documentEnd",
            "mode" => "plainText",
            "text" => "x",
            "cells" => "-",
            "widths" => "1000",
            "edges" => "top",
            "target" => "row",
            _ => throw new InvalidOperationException(
                $"{tool}.{parameter.Name} is a required string parameter with no baseline value in this sweep; add one."),
        };
    }

    private static IEnumerable<(string Tool, MethodInfo Method)> ToolMethods() =>
        typeof(ToolGuards).Assembly.GetTypes()
            .Where(t => t.GetCustomAttribute<McpServerToolTypeAttribute>() is not null)
            .SelectMany(t => t.GetMethods(BindingFlags.Public | BindingFlags.Static))
            .Select(m => (Tool: m.GetCustomAttribute<McpServerToolAttribute>()?.Name, Method: m))
            .Where(x => x.Tool is not null)
            .Select(x => (x.Tool!, x.Method))
            .OrderBy(x => x.Item1, StringComparer.Ordinal);

    // ---- unknown tool ----

    [Fact]
    public async Task An_unknown_tool_is_an_enveloped_invalid_argument_naming_it_not_a_protocol_error()
    {
        var result = await pipeline.CallAsync("no_such_tool", "{}");

        Assert.True(result.IsError);
        result.AssertFailure("invalid_argument", "'no_such_tool'");
        Assert.Contains("tools/list", result.Hint, StringComparison.Ordinal);
    }

    [Fact]
    public async Task An_unknown_tool_close_to_a_real_one_suggests_it()
    {
        var result = await pipeline.CallAsync("insert_feild", "{}");

        result.AssertFailure("invalid_argument", "'insert_feild'", "Did you mean 'insert_field'?");
    }

    // ---- unknown argument names (GitHub issue #90) ----

    [Theory]
    [InlineData("insert_text", """ "text":"x","locationType":"documentEnd","layoutpart":"footer" """, "layoutpart", "layoutPart")]
    [InlineData("insert_column", """ "tableIndex":0,"mode":"plainText","atcolumn":0 """, "atcolumn", "atColumn")]
    [InlineData("set_cell_format", """ "tableIndex":0,"row":0,"edges":"top","column":1 """, "column", "col")]
    [InlineData("insert_field", """ "field":"/Header/CustomerAddress1","locationType":"tableCell","tableIndex":0,"row":0,"col":0,"tabelIndex":1 """, "tabelIndex", "tableIndex")]
    public async Task An_unknown_argument_name_is_refused_naming_the_closest_real_parameter_and_nothing_is_written(
        string tool, string otherArguments, string sent, string meant)
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        var before = await File.ReadAllBytesAsync(path);

        var result = await pipeline.CallAsync(tool, $$"""{"layoutPath":{{J(path)}},{{otherArguments}}}""");

        Assert.True(result.IsError);
        result.AssertFailure("invalid_argument", $"no argument named '{sent}' (did you mean '{meant}'?)");
        Assert.Contains("case-sensitive", result.Hint, StringComparison.Ordinal);
        Assert.Contains($"{tool} takes: layoutPath, ", result.Hint, StringComparison.Ordinal);
        Assert.Equal(before, await File.ReadAllBytesAsync(path));
    }

    [Fact]
    public async Task An_unknown_argument_name_with_nothing_close_lists_the_valid_names_without_a_guess()
    {
        var result = await pipeline.CallAsync(
            "remove_control", $$"""{"layoutPath":{{J(MissingFile())}},"controlId":1,"zzz":true}""");

        result.AssertFailure("invalid_argument", "no argument named 'zzz';");
        Assert.DoesNotContain("did you mean", result.Message, StringComparison.Ordinal);
        Assert.Contains("remove_control takes: layoutPath, controlId, keepText.", result.Hint, StringComparison.Ordinal);
    }

    [Fact]
    public void ClosestName_is_conservative_rather_than_guessing()
    {
        string[] candidates = ["layoutPath", "tableIndex", "row", "col", "layoutPart"];

        Assert.Equal("layoutPart", ToolCallBoundary.ClosestName("LAYOUTPART", candidates));
        Assert.Equal("tableIndex", ToolCallBoundary.ClosestName("tableIdex", candidates));
        Assert.Equal("col", ToolCallBoundary.ClosestName("column", candidates));
        Assert.Null(ToolCallBoundary.ClosestName("foo", candidates));
        Assert.Null(ToolCallBoundary.ClosestName("x", candidates));
    }

    // ---- object arguments: format (GitHub issue #15) ----

    [Fact]
    public async Task A_format_sent_as_a_json_string_is_read_as_the_object_it_holds()
    {
        // Models do send a nested object JSON-encoded; the string is unambiguous, so it is read, not refused.
        var path = CopyOfCorpus(Corpus.SalesInvoice);

        var result = await pipeline.CallAsync(
            "insert_text",
            $$"""{"layoutPath":{{J(path)}},"text":"Formatted","locationType":"documentEnd","format":"{\"bold\":true,\"italic\":true}"}""");

        Assert.True(result.Ok, result.RawText);
        using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(path, false);
        var run = doc.MainDocumentPart!.Document!.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>()
            .Single(r => r.InnerText == "Formatted");
        Assert.NotNull(run.RunProperties?.Bold);
        Assert.NotNull(run.RunProperties?.Italic);
    }

    [Theory]
    [InlineData("""{"Bold":true}""", "'Bold' (did you mean 'bold'?)")]
    [InlineData("""{"fontsize":8}""", "'fontsize' (did you mean 'fontSizePoints'?)")]
    [InlineData("""{"bold":true,"strike":true}""", "'strike';")]
    public async Task An_unknown_format_member_is_refused_naming_the_closest_real_one_and_nothing_is_written(
        string format, string described)
    {
        // The binder drops an unknown member as silently as an unknown argument, so it is refused the same way.
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        var before = await File.ReadAllBytesAsync(path);

        var result = await pipeline.CallAsync(
            "insert_text", $$"""{"layoutPath":{{J(path)}},"text":"x","locationType":"documentEnd","format":{{format}}}""");

        Assert.True(result.IsError);
        result.AssertFailure("invalid_argument", $"insert_text's format has no member named {described}");
        Assert.Contains(
            "format takes: bold, italic, underline, fontSizePoints, fontName, color, characterStyle.",
            result.Hint,
            StringComparison.Ordinal);
        Assert.Equal(before, await File.ReadAllBytesAsync(path));
    }

    [Theory]
    [InlineData("""{"bold":"yes"}""", "format.bold", "a boolean (true or false) or null", "the string \"yes\"")]
    [InlineData("""{"fontSizePoints":true}""", "format.fontSizePoints", "a number or null", "the boolean true")]
    public async Task A_wrong_typed_format_member_is_invalid_argument_naming_the_member(
        string format, string member, string expected, string sent)
    {
        var result = await pipeline.CallAsync(
            "insert_text", $$"""{"layoutPath":{{J(MissingFile())}},"text":"x","locationType":"documentEnd","format":{{format}}}""");

        Assert.True(result.IsError);
        result.AssertFailure("invalid_argument", $"Member '{member}' of insert_text", $"must be {expected}", $"sent {sent}");
    }

    [Theory]
    [InlineData("5", "the number 5")]
    [InlineData("\"bold\"", "the string \"bold\"")]
    public async Task A_format_that_is_not_an_object_is_invalid_argument_asking_for_one(string format, string sent)
    {
        var result = await pipeline.CallAsync(
            "insert_text", $$"""{"layoutPath":{{J(MissingFile())}},"text":"x","locationType":"documentEnd","format":{{format}}}""");

        Assert.True(result.IsError);
        result.AssertFailure("invalid_argument", "Argument 'format' of insert_text", "must be a JSON object", $"sent {sent}");
    }

    // ---- relative paths (GitHub issue #90) ----

    [Fact]
    public async Task Relative_paths_are_refused_naming_each_parameter_and_nothing_is_written_under_the_working_directory()
    {
        var outputUnderCwd = Path.Combine(Environment.CurrentDirectory, "sub", "New.docx");

        var result = await pipeline.CallAsync(
            "create_layout", """{"schemaSource":"RelQuote.docx","outputPath":"sub/New.docx"}""");

        Assert.True(result.IsError);
        result.AssertFailure("invalid_argument", "schemaSource ('RelQuote.docx')", "outputPath ('sub/New.docx')", "relative");
        Assert.Contains("fully qualified", result.Hint, StringComparison.Ordinal);
        Assert.False(File.Exists(outputUnderCwd));
    }

    [Theory]
    [InlineData("get_layout_info", "layoutPath", "Layout.docx")]
    [InlineData("get_layout_info", "layoutPath", @"..\Layout.docx")]
    [InlineData("get_layout_info", "layoutPath", @"\Layout.docx")]
    [InlineData("get_layout_info", "layoutPath", "C:Layout.docx")]
    [InlineData("list_dataset_fields", "source", "schema.xml")]
    [InlineData("render_preview_pages", "pdfPath", "preview.pdf")]
    public async Task A_relative_or_drive_or_root_relative_path_is_refused(string tool, string parameter, string path)
    {
        var result = await pipeline.CallAsync(tool, $$"""{"{{parameter}}":{{J(path)}}}""");

        result.AssertFailure("invalid_argument", $"{parameter} ('{path}')");
    }

    [Fact]
    public async Task A_relative_optional_output_directory_is_refused_before_any_preview_work()
    {
        // preview_layout's own file check would also stop this call (the layout is missing); the point is
        // that the relative outputDir is what gets named, before the tool body runs at all.
        var result = await pipeline.CallAsync(
            "preview_layout", $$"""{"layoutPath":{{J(MissingFile())}},"outputDir":"previews"}""");

        result.AssertFailure("invalid_argument", "outputDir ('previews')");
    }

    [Fact]
    public async Task A_blank_path_passes_through_to_the_tools_own_more_precise_refusal()
    {
        var result = await pipeline.CallAsync(
            "create_layout", $$"""{"schemaSource":{{J(CopyOfCorpus(Corpus.SalesInvoice))}},"outputPath":""}""");

        result.AssertFailure("invalid_argument", "Output path must not be empty.");
    }

    /// <summary>
    /// The relative-path check only covers parameters in <c>ToolCallBoundary.FilePathParameters</c>, so a new
    /// path parameter missing from that set would silently accept relative paths again. Every string parameter
    /// that LOOKS like a file path — by name, or by a description promising an absolute path — must be either in
    /// the set or explicitly classified here as a DATASET path (an XPath-like '/Header/Line/...' string, which
    /// is never a file); and every entry in the set must still be a real string parameter of some tool.
    /// </summary>
    [Fact]
    public void Every_path_like_string_parameter_is_classified_as_a_file_path_or_a_dataset_path()
    {
        string[] datasetPathParameters = ["dataPath", "headerLabelPath"];
        string[] pathLikeSuffixes = ["path", "source", "dir", "directory", "file", "folder"];

        var stringParameters = ToolMethods()
            .SelectMany(x => x.Method.GetParameters()
                .Where(p => p.ParameterType == typeof(string))
                .Select(p => (x.Tool, Parameter: p)))
            .ToList();

        var unclassified = stringParameters
            .Where(x => pathLikeSuffixes.Any(s => x.Parameter.Name!.EndsWith(s, StringComparison.OrdinalIgnoreCase))
                        || (x.Parameter.GetCustomAttribute<System.ComponentModel.DescriptionAttribute>()?.Description
                            .Contains("absolute", StringComparison.OrdinalIgnoreCase) ?? false))
            .Where(x => !ToolCallBoundary.FilePathParameters.Contains(x.Parameter.Name!)
                        && !datasetPathParameters.Contains(x.Parameter.Name))
            .Select(x => $"{x.Tool}.{x.Parameter.Name}")
            .ToList();
        Assert.True(unclassified.Count == 0,
            "Path-like string parameters that are neither in ToolCallBoundary.FilePathParameters nor classified as "
            + $"dataset paths: {string.Join(", ", unclassified)}.");

        var stale = ToolCallBoundary.FilePathParameters
            .Where(name => !stringParameters.Any(x => x.Parameter.Name == name))
            .ToList();
        Assert.True(stale.Count == 0, $"FilePathParameters names no tool's string parameter: {string.Join(", ", stale)}.");
    }

    // ---- numeric ranges (GitHub issue #98) ----

    [Theory]
    [InlineData("set_column_widths", """ "tableIndex":0,"widths":"-500,5103" """, "widths entry '-500'", "set_column_widths' widths")]
    [InlineData("set_column_widths", """ "tableIndex":0,"widths":"0,0" """, "widths entry '0'", "set_column_widths' widths")]
    [InlineData("set_column_widths", """ "tableIndex":0,"widths":"40000,100" """, "widths entry '40000'", "set_column_widths' widths")]
    [InlineData("insert_table", """ "rows":1,"columns":2,"locationType":"documentEnd","columnWidths":"-100,200" """, "columnWidths entry '-100'", "columnWidths, if supplied")]
    [InlineData("insert_table", """ "rows":1,"columns":2,"locationType":"documentEnd","columnWidths":"0,200" """, "columnWidths entry '0'", "columnWidths, if supplied")]
    [InlineData("insert_repeater_table", """ "dataItem":"/Header/Line","columns":"ItemNo_Line","locationType":"documentEnd","columnWidths":"-5" """, "columnWidths entry '-5'", "columnWidths, if supplied")]
    [InlineData("insert_column", """ "tableIndex":0,"mode":"plainText","width":40000 """, "width 40000", "insert_column's width")]
    [InlineData("insert_column", """ "tableIndex":0,"mode":"plainText","width":0 """, "width 0", "insert_column's width")]
    public async Task A_column_width_outside_1_to_Words_maximum_page_width_is_invalid_argument_and_nothing_is_written(
        string tool, string otherArguments, string messageFragment, string hintFragment)
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        var before = await File.ReadAllBytesAsync(path);

        var result = await pipeline.CallAsync(tool, $$"""{"layoutPath":{{J(path)}},{{otherArguments}}}""");

        // Was edit_would_corrupt ("likely a bug in the tool") for a negative width and ok:true - writing
        // w:gridCol w:w="0" - for a zero one.
        result.AssertFailure("invalid_argument", messageFragment);
        Assert.Contains(hintFragment, result.Hint, StringComparison.Ordinal);
        Assert.Equal(before, await File.ReadAllBytesAsync(path));
    }

    [Fact]
    public async Task Column_widths_at_both_bounds_are_accepted()
    {
        var result = await pipeline.CallAsync(
            "set_column_widths", $$"""{"layoutPath":{{J(CopyOfCorpus(Corpus.SalesInvoice))}},"tableIndex":0,"widths":"1,{{TableTools.MaxColumnWidthTwips}}"}""");

        Assert.True(result.Ok, result.RawText);
    }

    [Theory]
    [InlineData("insert_table_row", """ "tableIndex":0,"cells":"3:-,-1:-" """, "'-1:-' spans -1")]
    [InlineData("insert_table_row", """ "tableIndex":0,"cells":"-5:-,13:-" """, "'-5:-' spans -5")]
    [InlineData("insert_table_row", """ "tableIndex":0,"cells":"0:-,2:-" """, "'0:-' spans 0")]
    [InlineData("insert_table_row", """ "tableIndex":0,"cells":"64:-" """, "'64:-' spans 64")]
    [InlineData("insert_subtotal_row", """ "parentControlId":0,"cells":"1:-,-1:-" """, "'-1:-' spans -1")]
    [InlineData("insert_repeater_row", """ "parentControlId":0,"dataItem":"/Header/Line","cells":"-3:-" """, "'-3:-' spans -3")]
    public async Task A_cell_span_outside_1_to_63_is_invalid_argument_with_the_cells_hint(
        string tool, string otherArguments, string messageFragment)
    {
        var result = await pipeline.CallAsync(
            tool, $$"""{"layoutPath":{{J(CopyOfCorpus(Corpus.SalesInvoice))}},{{otherArguments}}}""");

        // "3:-,-1:-" used to crash with "Index was out of range (Parameter 'index')" and a generic hint.
        result.AssertFailure("invalid_argument", messageFragment, "between 1 and 63");
        Assert.Contains("cells is a comma-separated list of cell specs", result.Hint, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(""" "rows":1,"columns":0 """, "columns must be between 1 and 30", "insert_table's columns is how many columns")]
    [InlineData(""" "rows":1,"columns":31 """, "columns must be between 1 and 30", "insert_table's columns is how many columns")]
    [InlineData(""" "rows":0,"columns":2 """, "rows must be between 1 and 100", "insert_table's rows is how many rows")]
    public async Task Insert_table_count_hints_describe_insert_tables_own_parameters(
        string counts, string messageFragment, string hintFragment)
    {
        var result = await pipeline.CallAsync(
            "insert_table", $$"""{"layoutPath":{{J(CopyOfCorpus(Corpus.SalesInvoice))}},{{counts}},"locationType":"documentEnd"}""");

        // The columns hint used to describe insert_repeater_table's comma-separated list of column names.
        result.AssertFailure("invalid_argument", messageFragment);
        Assert.Contains(hintFragment, result.Hint, StringComparison.Ordinal);
        Assert.DoesNotContain("LEAF column names", result.Hint, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Insert_repeater_tables_columns_keeps_its_own_list_hint()
    {
        var result = await pipeline.CallAsync(
            "insert_repeater_table", $$"""{"layoutPath":{{J(CopyOfCorpus(Corpus.SalesInvoice))}},"dataItem":"/Header/Line","columns":" , ","locationType":"documentEnd"}""");

        result.AssertFailure("invalid_argument", "At least one column is required");
        Assert.Contains("insert_repeater_table's columns must be a non-empty, comma-separated list", result.Hint, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-5)]
    public async Task Preview_rows_below_1_is_invalid_argument_with_the_preview_hint(int rows)
    {
        // The layout is missing, so even a regression in the rows check could not reach a conversion: the
        // call would stop at file_not_found instead, and fail this test.
        var result = await pipeline.CallAsync(
            "preview_layout", $$"""{"layoutPath":{{J(MissingFile())}},"rows":{{rows}}}""");

        result.AssertFailure("invalid_argument", $"rows must be 1 or more (got {rows})");
        Assert.Contains("preview_layout's rows is how many sample rows", result.Hint, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("firstPage", 0)]
    [InlineData("firstPage", -3)]
    [InlineData("maxPages", 0)]
    public async Task Render_page_arguments_below_1_are_invalid_argument_instead_of_a_silent_clamp(string parameter, int value)
    {
        // Was: page 1 rendered while the response echoed firstPage 0/-3. The PDF is missing, so nothing renders
        // either way; the argument check runs first.
        var result = await pipeline.CallAsync(
            "render_preview_pages", $$"""{"pdfPath":{{J(MissingFile(".pdf"))}},"{{parameter}}":{{value}}}""");

        result.AssertFailure("invalid_argument", $"{parameter} must be 1 or more (got {value})");
        Assert.Contains("pageCount", result.Hint, StringComparison.Ordinal);
    }

    // ---- belt-and-braces: an SDK-swallowed binding failure ----

    [Fact]
    public async Task A_non_envelope_isError_result_is_rewrapped_as_invalid_argument_never_an_undocumented_code()
    {
        var result = await pipeline.CallAsync(SwallowedFailureTestTool.Name, "{}");

        Assert.True(result.IsError);
        result.AssertFailure("invalid_argument", SwallowedFailureTestTool.SdkText);
    }
}
