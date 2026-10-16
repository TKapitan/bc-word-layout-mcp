using System.Reflection;
using System.Text.Json;
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
    [InlineData("set_cell_borders", """ "tableIndex":0,"row":0,"edges":"top","size":null """, "size", "an integer", "null")]
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
    [InlineData("insert_repeater_table", """ "dataItem":"/Header/Line","columns":"ItemNo_Line","locationType":"documentEnd","look":null """)]
    [InlineData("set_cell_borders", """ "tableIndex":0,"row":0,"edges":"top","style":null """)]
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

    // ---- belt-and-braces: an SDK-swallowed binding failure ----

    [Fact]
    public async Task A_non_envelope_isError_result_is_rewrapped_as_invalid_argument_never_an_undocumented_code()
    {
        var result = await pipeline.CallAsync(SwallowedFailureTestTool.Name, "{}");

        Assert.True(result.IsError);
        result.AssertFailure("invalid_argument", SwallowedFailureTestTool.SdkText);
    }
}
