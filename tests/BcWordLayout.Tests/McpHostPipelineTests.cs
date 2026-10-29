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

        // Nullable since SDK 2.3.0: a server may send no identity at all (GitHub issue #153). Ours must send one.
        var serverInfo = pipeline.Client.ServerInfo;
        Assert.NotNull(serverInfo);

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
    [InlineData("insert_table", """ "rowCount":1.5,"columns":[{},{}],"locationType":"documentEnd" """, "rowCount", "an integer", "the number 1.5")]
    [InlineData("set_cell_format", """ "tableIndex":0,"row":0,"edges":["top"],"borderSize":"thick" """, "borderSize", "an integer or null", "the string \"thick\"")]
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
    [InlineData("insert_repeater_table", """ "dataItem":"/Header/Line","columns":[{"field":"/Header/Line/ItemNo_Line"}],"locationType":"documentEnd","borders":null """)]
    [InlineData("insert_table", """ "rowCount":1,"columns":[{},{}],"locationType":"documentEnd","borders":null """)]
    [InlineData("set_cell_format", """ "tableIndex":0,"row":0,"edges":["top"],"borderStyle":null """)]
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
    /// Sends an explicit JSON null for EVERY parameter of EVERY tool that can bind one — strings, and the object and
    /// array arguments (GitHub issue #133) — one at a time, with the other required arguments set to plausible
    /// values: none may come back as <c>internal_error</c> (the "bug in the tool" code). New tools and parameters
    /// are covered automatically; a new REQUIRED parameter name needs a baseline value below, and the test says
    /// so. <c>preview_layout</c> is always pointed at a missing layout and the converter seam at a fake that
    /// records any use, so the sweep can never start a conversion.
    /// </summary>
    [Fact]
    public async Task No_parameter_of_any_tool_turns_an_explicit_null_into_internal_error()
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
                foreach (var nulled in method.GetParameters().Where(p => !p.ParameterType.IsValueType))
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
                "rowCount" or "toColumn" => 1,
                _ => 0,
            };
        }

        if (parameter.ParameterType == typeof(IReadOnlyList<int>))
        {
            return new[] { 1000 };
        }

        if (parameter.ParameterType == typeof(IReadOnlyList<RepeaterColumnArgument>))
        {
            return new[] { new { field = "/Header/Line/ItemNo_Line" } };
        }

        if (parameter.ParameterType == typeof(IReadOnlyList<TableColumnArgument>))
        {
            return new[] { new { width = 3000 } };
        }

        if (parameter.ParameterType == typeof(IReadOnlyList<RowCellArgument>))
        {
            return new[] { new { span = 1 } };
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
            "locationType" => "documentEnd",
            "mode" => "plainText",
            "text" => "x",
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
    [InlineData("set_cell_format", """ "tableIndex":0,"row":0,"edges":["top"],"column":1 """, "column", "col")]
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

    [Theory]
    [InlineData("layoutPart", "header")]
    [InlineData("partName", "header2.xml")]
    public async Task Insert_repeater_table_has_no_part_arguments_so_a_header_target_is_refused_as_unknown(
        string argument, string value)
    {
        // A repeater table is body-only, and the tool used to take layoutPart/partName anyway, described as "always
        // omit" (GitHub issue #132). Without them it is body-only by construction: the refusal names the argument
        // like any other unknown one, and nothing is written.
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        var before = await File.ReadAllBytesAsync(path);

        var result = await pipeline.CallAsync(
            "insert_repeater_table",
            $$"""{"layoutPath":{{J(path)}},"dataItem":"/Header/Line","columns":[{"field":"/Header/Line/ItemNo_Line"}],"locationType":"documentEnd","{{argument}}":"{{value}}"}""");

        result.AssertFailure("invalid_argument", $"insert_repeater_table has no argument named '{argument}'");
        Assert.Equal(before, await File.ReadAllBytesAsync(path));
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

    // ---- array arguments (GitHub issue #133) ----

    [Fact]
    public async Task A_list_sent_as_a_json_string_is_read_as_the_array_it_holds()
    {
        // The same tolerance as for format: the string is unambiguous JSON, so it is read rather than refused.
        var path = CopyOfCorpus(Corpus.SalesInvoice);

        var result = await pipeline.CallAsync(
            "set_column_widths", $$"""{"layoutPath":{{J(path)}},"tableIndex":0,"widths":"[1500, 8706]"}""");

        Assert.True(result.Ok, result.RawText);
    }

    [Fact]
    public async Task A_comma_separated_list_is_invalid_argument_asking_for_a_json_array()
    {
        // The pre-2.0 form. It is not read as a list: a comma could never be part of a value, and empty entries
        // were silently dropped.
        var result = await pipeline.CallAsync(
            "set_column_widths", $$"""{"layoutPath":{{J(MissingFile())}},"tableIndex":0,"widths":"1500,8706"}""");

        result.AssertFailure(
            "invalid_argument", "Argument 'widths' of set_column_widths", "must be a JSON array of integers", "sent the string \"1500,8706\"");
        Assert.Contains("a list is a JSON array ([1500, 6000]), not a comma-separated string", result.Hint, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("set_column_widths", """ "tableIndex":0,"widths":[1500,"wide"] """, "Item 'widths[1]'", "an integer", "the string \"wide\"")]
    [InlineData("set_column_widths", """ "tableIndex":0,"widths":[1500,null] """, "Item 'widths[1]'", "an integer", "null")]
    [InlineData("set_cell_format", """ "tableIndex":0,"row":0,"edges":["top",5] """, "Item 'edges[1]'", "a string", "the number 5")]
    [InlineData("set_cell_format", """ "tableIndex":0,"row":0,"edges":"top" """, "Argument 'edges'", "a JSON array of strings or null", "the string \"top\"")]
    public async Task A_wrong_typed_list_or_item_is_invalid_argument_naming_exactly_that_item(
        string tool, string otherArguments, string what, string expected, string sent)
    {
        var result = await pipeline.CallAsync(tool, $$"""{"layoutPath":{{J(MissingFile())}},{{otherArguments}}}""");

        Assert.True(result.IsError);
        result.AssertFailure("invalid_argument", $"{what} of {tool}", $"must be {expected},", $"sent {sent}");
    }

    [Fact]
    public async Task The_input_schema_puts_a_lists_allowed_values_and_range_on_its_items()
    {
        // The SDK writes [AllowedValues]/[Range] on a LIST parameter onto the array node, where no value can satisfy
        // them; ToolInputSchemas moves them onto items, so tools/list says what each entry may be.
        var tools = await pipeline.Client.ListToolsAsync();
        JsonElement Parameter(string tool, string name) =>
            tools.Single(t => t.Name == tool).ProtocolTool.InputSchema.GetProperty("properties").GetProperty(name);

        var edges = Parameter("set_cell_format", "edges");
        Assert.False(edges.TryGetProperty("enum", out _));
        Assert.Equal(
            new[] { "top", "bottom", "left", "right", "all" },
            edges.GetProperty("items").GetProperty("enum").EnumerateArray().Select(e => e.GetString()));
        Assert.Equal(1, edges.GetProperty("minItems").GetInt32());

        var widths = Parameter("set_column_widths", "widths");
        Assert.False(widths.TryGetProperty("minimum", out _));
        Assert.Equal(1, widths.GetProperty("items").GetProperty("minimum").GetInt32());
        Assert.Equal(TableTools.MaxColumnWidthTwips, widths.GetProperty("items").GetProperty("maximum").GetInt32());
        Assert.Equal("integer", widths.GetProperty("items").GetProperty("type").GetString());
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
    [InlineData("set_column_widths", """ "tableIndex":0,"widths":[-500,5103] """, "widths[0] (-500)", "set_column_widths' widths")]
    [InlineData("set_column_widths", """ "tableIndex":0,"widths":[5103,0] """, "widths[1] (0)", "set_column_widths' widths")]
    [InlineData("set_column_widths", """ "tableIndex":0,"widths":[40000,100] """, "widths[0] (40000)", "set_column_widths' widths")]
    [InlineData("insert_table", """ "rowCount":1,"columns":[{"width":-100},{"width":200}],"locationType":"documentEnd" """, "columns[0].width (-100)", "A width is in twips")]
    [InlineData("insert_table", """ "rowCount":1,"columns":[{"width":200},{"width":0}],"locationType":"documentEnd" """, "columns[1].width (0)", "A width is in twips")]
    [InlineData("insert_repeater_table", """ "dataItem":"/Header/Line","columns":[{"field":"/Header/Line/ItemNo_Line","width":-5}],"locationType":"documentEnd" """, "columns[0].width (-5)", "A width is in twips")]
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
            "set_column_widths", $$"""{"layoutPath":{{J(CopyOfCorpus(Corpus.SalesInvoice))}},"tableIndex":0,"widths":[1,{{TableTools.MaxColumnWidthTwips}}]}""");

        Assert.True(result.Ok, result.RawText);
    }

    [Theory]
    [InlineData("insert_table_row", """ "tableIndex":0,"cells":[{"span":3},{"span":-1}] """, "cells[1].span is -1")]
    [InlineData("insert_table_row", """ "tableIndex":0,"cells":[{"span":-5},{"span":13}] """, "cells[0].span is -5")]
    [InlineData("insert_table_row", """ "tableIndex":0,"cells":[{"span":0},{"span":2}] """, "cells[0].span is 0")]
    [InlineData("insert_table_row", """ "tableIndex":0,"cells":[{"span":64}] """, "cells[0].span is 64")]
    [InlineData("insert_subtotal_row", """ "parentControlId":0,"cells":[{"span":1},{"span":-1}] """, "cells[1].span is -1")]
    [InlineData("insert_repeater_row", """ "parentControlId":0,"dataItem":"/Header/Line","cells":[{"span":-3}] """, "cells[0].span is -3")]
    public async Task A_cell_span_outside_1_to_63_is_invalid_argument_with_the_cells_hint(
        string tool, string otherArguments, string messageFragment)
    {
        var result = await pipeline.CallAsync(
            tool, $$"""{"layoutPath":{{J(CopyOfCorpus(Corpus.SalesInvoice))}},{{otherArguments}}}""");

        // A span of -1 used to crash with "Index was out of range (Parameter 'index')" and a generic hint.
        result.AssertFailure("invalid_argument", messageFragment, "between 1 and 63");
        Assert.Contains("cells is an array with one object per cell", result.Hint, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("insert_table", """ "rowCount":1,"columns":[] """, "columns must hold at least one entry", "columns is an array with one object per table column")]
    [InlineData("insert_table", """ "rowCount":1,"columns":[{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{}] """, "columns must hold between 1 and 30 columns (got 31)", "columns is an array with one object per table column")]
    [InlineData("insert_table", """ "rowCount":0,"columns":[{},{}] """, "rowCount must be between 1 and 100", "insert_table's rowCount is how many rows")]
    [InlineData("insert_repeater_table", """ "dataItem":"/Header/Line","columns":[] """, "columns must hold at least one entry", "columns is an array with one object per table column")]
    public async Task Insert_table_count_hints_describe_the_tools_own_parameters(
        string tool, string counts, string messageFragment, string hintFragment)
    {
        // insert_table's columns was a count while insert_repeater_table's was a list of names, and rows meant
        // something else on preview_layout, so their hints had to be keyed per tool. Both tables now take the same
        // array of column objects, and the counts are rowCount and sampleRows (GitHub issue #133).
        var result = await pipeline.CallAsync(
            tool, $$"""{"layoutPath":{{J(CopyOfCorpus(Corpus.SalesInvoice))}},{{counts}},"locationType":"documentEnd"}""");

        result.AssertFailure("invalid_argument", messageFragment);
        Assert.Contains(hintFragment, result.Hint, StringComparison.Ordinal);
    }

    [Fact]
    public async Task An_unknown_member_of_an_object_in_an_array_is_refused_naming_its_item()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        var before = await File.ReadAllBytesAsync(path);

        var result = await pipeline.CallAsync(
            "insert_table", $$"""{"layoutPath":{{J(path)}},"rowCount":1,"columns":[{},{"Width":1500}],"locationType":"documentEnd"}""");

        Assert.True(result.IsError);
        result.AssertFailure("invalid_argument", "insert_table's columns[1] has no member named 'Width' (did you mean 'width'?)");
        Assert.Contains("Each item of columns takes: width, alignment.", result.Hint, StringComparison.Ordinal);
        Assert.Equal(before, await File.ReadAllBytesAsync(path));
    }

    [Fact]
    public async Task An_object_in_an_array_without_a_required_member_is_refused_naming_it()
    {
        var result = await pipeline.CallAsync(
            "insert_repeater_table",
            $$"""{"layoutPath":{{J(MissingFile())}},"dataItem":"/Header/Line","columns":[{"field":"/Header/Line/ItemNo_Line"},{"width":1500}],"locationType":"documentEnd"}""");

        Assert.True(result.IsError);
        result.AssertFailure("invalid_argument", "'columns[1]' of insert_repeater_table has no 'field', which it requires.");
    }

    [Theory]
    [InlineData("insert_table", """ "rowCount":1,"columns":[{"width":"wide"}],"locationType":"documentEnd" """, "Member 'columns[0].width'", "an integer or null", "the string \"wide\"")]
    [InlineData("insert_table_row", """ "tableIndex":0,"cells":[{"span":2,"fields":"/Header/No"}] """, "Member 'cells[0].fields'", "a JSON array of strings or null", "the string \"/Header/No\"")]
    [InlineData("insert_table_row", """ "tableIndex":0,"cells":[5] """, "Item 'cells[0]'", "a JSON object", "the number 5")]
    public async Task A_wrong_typed_member_of_an_object_in_an_array_is_named_by_its_path(
        string tool, string otherArguments, string what, string expected, string sent)
    {
        var result = await pipeline.CallAsync(tool, $$"""{"layoutPath":{{J(MissingFile())}},{{otherArguments}}}""");

        Assert.True(result.IsError);
        result.AssertFailure("invalid_argument", $"{what} of {tool}", $"must be {expected}", $"sent {sent}");
    }

    [Fact]
    public async Task An_array_of_objects_sent_as_a_json_string_is_read_as_the_array_it_holds()
    {
        // SalesInvoice's first table has two grid columns: one cell spanning both covers it.
        var path = CopyOfCorpus(Corpus.SalesInvoice);

        var result = await pipeline.CallAsync(
            "insert_table_row", $$"""{"layoutPath":{{J(path)}},"tableIndex":0,"cells":"[{\"span\": 2}]"}""");

        Assert.True(result.Ok, result.RawText);
    }

    [Fact]
    public async Task The_input_schema_lists_the_values_of_enumerated_arguments_and_the_range_of_bounded_numbers()
    {
        // The value lists of AllowedValueAttributes.cs reach tools/list as "enum" through the SDK, and [Range] as
        // minimum/maximum, so an agent reads the choices from the schema rather than from prose.
        var tools = await pipeline.Client.ListToolsAsync();
        JsonElement Parameter(string tool, string name) =>
            tools.Single(t => t.Name == tool).ProtocolTool.InputSchema.GetProperty("properties").GetProperty(name);
        string?[] Values(JsonElement schema) =>
            [.. schema.GetProperty("enum").EnumerateArray().Select(v => v.ValueKind == JsonValueKind.Null ? null : v.GetString())];

        Assert.Equal(new[] { "documentEnd", "afterControl", "tableCell", "atText" }, Values(Parameter("insert_field", "locationType")));
        Assert.Equal(new[] { "body", "header", "footer" }, Values(Parameter("set_cell_format", "layoutPart")));
        Assert.Equal(new[] { "left", "center", "right" }, Values(Parameter("set_cell_text", "alignment")));
        Assert.Equal(new[] { "none", "headerRule", "grid" }, Values(Parameter("insert_repeater_table", "borders")));
        Assert.Equal(new[] { "quick", "full" }, Values(Parameter("validate_layout", "level")));

        var borderSize = Parameter("set_cell_format", "borderSize");
        Assert.Equal(2, borderSize.GetProperty("minimum").GetInt32());
        Assert.Equal(96, borderSize.GetProperty("maximum").GetInt32());
        var fontSize = Parameter("insert_text", "format").GetProperty("properties").GetProperty("fontSizePoints");
        Assert.Equal(4, fontSize.GetProperty("minimum").GetDouble());
        Assert.Equal(96, fontSize.GetProperty("maximum").GetDouble());
    }

    [Fact]
    public async Task The_input_schema_describes_each_column_and_cell_object()
    {
        // One object per column or cell (GitHub issue #133): tools/list says what each member may hold.
        var tools = await pipeline.Client.ListToolsAsync();
        JsonElement Items(string tool, string parameter) =>
            tools.Single(t => t.Name == tool).ProtocolTool.InputSchema.GetProperty("properties").GetProperty(parameter).GetProperty("items");

        var repeaterColumn = Items("insert_repeater_table", "columns");
        Assert.Equal("field", Assert.Single(repeaterColumn.GetProperty("required").EnumerateArray()).GetString());
        Assert.Equal(TableTools.MaxColumnWidthTwips, repeaterColumn.GetProperty("properties").GetProperty("width").GetProperty("maximum").GetInt32());

        var tableColumn = Items("insert_table", "columns").GetProperty("properties");
        Assert.Equal(new[] { "width", "alignment" }, tableColumn.EnumerateObject().Select(p => p.Name));
        Assert.Equal(
            new[] { "left", "center", "right" },
            tableColumn.GetProperty("alignment").GetProperty("enum").EnumerateArray()
                .Where(v => v.ValueKind == JsonValueKind.String).Select(v => v.GetString()));

        var cell = Items("insert_table_row", "cells").GetProperty("properties");
        Assert.Equal(1, cell.GetProperty("span").GetProperty("minimum").GetInt32());
        Assert.Equal(TableTools.MaxCellSpan, cell.GetProperty("span").GetProperty("maximum").GetInt32());
        Assert.Equal("array", cell.GetProperty("fields").GetProperty("type").EnumerateArray().First().GetString());
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-5)]
    public async Task Preview_sampleRows_below_1_is_invalid_argument_with_the_preview_hint(int sampleRows)
    {
        // The layout is missing, so even a regression in the sampleRows check could not reach a conversion: the
        // call would stop at file_not_found instead, and fail this test.
        var result = await pipeline.CallAsync(
            "preview_layout", $$"""{"layoutPath":{{J(MissingFile())}},"sampleRows":{{sampleRows}}}""");

        result.AssertFailure("invalid_argument", $"sampleRows must be 1 or more (got {sampleRows})");
        Assert.Contains("preview_layout's sampleRows is how many sample rows", result.Hint, StringComparison.Ordinal);
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
