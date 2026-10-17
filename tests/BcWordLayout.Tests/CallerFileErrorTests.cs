using System.IO.Compression;
using System.Text;
using BcWordLayout.Domain;
using BcWordLayout.McpHost;
using BcWordLayout.McpHost.Tools;
using BcWordLayout.Render;

namespace BcWordLayout.Tests;

/// <summary>
/// GitHub issue #95: a caller-supplied file that is not what its parameter takes must be reported as the
/// caller's input problem, naming THAT file, never as <c>internal_error</c> ("likely a bug in the tool"):
/// <list type="bullet">
/// <item>Not a <c>.docx</c> package at all (an empty file, a text file or a schema <c>.xml</c> named
/// <c>.docx</c>, a truncated copy), through every tool that opens a package → <c>invalid_layout</c>, with a
/// hint that names the file and explains <c>.docx</c> versus schema <c>.xml</c>. The edit and preview tools
/// open a PRIVATE copy, so the message must name the caller's path, not the copy's.</item>
/// <item>A schema <c>.xml</c> that is not well-formed (truncated, a DTD, a wrong encoding declaration, a
/// <c>.docx</c> under an <c>.xml</c> name) → <c>invalid_layout</c>, naming the file and the line/position.</item>
/// <item>A <c>.docx</c> with a damaged XML part, parsed lazily by the SDK → <c>invalid_layout</c> naming the
/// layout (or, for <c>create_layout</c>, the template).</item>
/// <item>A <c>dataOverridesPath</c> file that cannot be used, for any reason → <c>invalid_argument</c> for
/// that parameter, with a hint listing the two accepted shapes.</item>
/// <item><see cref="LayoutBuilder"/>'s internal-bug gate keeps reporting <c>internal_error</c>.</item>
/// </list>
/// Every input is written into a private directory; the shared corpus is only ever copied.
/// </summary>
/// <remarks>Joins the preview-converter-seam collection because some tests call
/// <c>LifecycleTools.PreviewLayout</c> with a <see cref="FakePdfConverter"/> (see
/// <see cref="PreviewConverterSeamCollection"/>), which no failing merge ever reaches.</remarks>
[Collection("preview-converter-seam")]
public class CallerFileErrorTests
{
    private const string MinimalSchemaXml =
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
        + "<NavWordReportXmlPart xmlns=\"urn:microsoft-dynamics-nav/reports/Test/1/\">"
        + "<Header><CompanyName>X</CompanyName></Header>"
        + "</NavWordReportXmlPart>";

    private static string PrivateDirectory() => Directory.CreateTempSubdirectory("callerfile-").FullName;

    private static string CopyOfCorpusInto(string dir, string corpusFile, string fileName)
    {
        var path = Path.Combine(dir, fileName);
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    private static string WriteSchemaXml(string dir)
    {
        var path = Path.Combine(dir, "schema.xml");
        File.WriteAllText(path, MinimalSchemaXml, new UTF8Encoding(false));
        return path;
    }

    /// <summary>Writes a file named <c>.docx</c> that is not a Word package, in the shape named by <paramref name="kind"/>.</summary>
    private static string WriteNonPackage(string dir, string kind)
    {
        var path = Path.Combine(dir, $"not-a-package-{kind}.docx");
        switch (kind)
        {
            case "empty":
                File.WriteAllBytes(path, Array.Empty<byte>());
                break;
            case "text":
                File.WriteAllText(path, "This is a plain text file someone saved with a .docx name.");
                break;
            case "schema-xml":
                File.WriteAllText(path, MinimalSchemaXml, new UTF8Encoding(false));
                break;
            case "truncated":
                var bytes = File.ReadAllBytes(Corpus.Path(Corpus.SalesInvoice));
                File.WriteAllBytes(path, bytes.AsSpan(0, bytes.Length / 2).ToArray());
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(kind), kind, null);
        }

        return path;
    }

    /// <summary>Writes a schema <c>.xml</c> that is not well-formed XML, in the shape named by <paramref name="kind"/>.</summary>
    private static string WriteMalformedSchemaXml(string dir, string kind)
    {
        var path = Path.Combine(dir, $"malformed-{kind}.xml");
        switch (kind)
        {
            case "truncated":
                File.WriteAllText(path, MinimalSchemaXml[..(MinimalSchemaXml.Length - 30)], new UTF8Encoding(false));
                break;
            case "dtd":
                File.WriteAllText(
                    path,
                    "<?xml version=\"1.0\"?><!DOCTYPE NavWordReportXmlPart [<!ENTITY e \"x\">]>"
                    + "<NavWordReportXmlPart xmlns=\"urn:microsoft-dynamics-nav/reports/Test/1/\"><Header><A>&e;</A></Header></NavWordReportXmlPart>",
                    new UTF8Encoding(false));
                break;
            case "utf16-declared-on-utf8":
                File.WriteAllText(path, MinimalSchemaXml.Replace("utf-8", "utf-16", StringComparison.Ordinal), new UTF8Encoding(false));
                break;
            case "docx-bytes":
                File.Copy(Corpus.Path(Corpus.SalesInvoice), path, overwrite: true);
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(kind), kind, null);
        }

        return path;
    }

    /// <summary>A copy of a corpus layout whose main document part is truncated mid-element (a damaged package).</summary>
    private static string WriteDamagedPackage(string dir, string fileName)
    {
        var path = CopyOfCorpusInto(dir, Corpus.SalesInvoice, fileName);
        using var zip = ZipFile.Open(path, ZipArchiveMode.Update);
        zip.GetEntry("word/document.xml")!.Delete();
        using var writer = new StreamWriter(zip.CreateEntry("word/document.xml").Open());
        writer.Write("<?xml version=\"1.0\"?><w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p>");
        return path;
    }

    private static void AssertInvalidLayoutNaming(ToolResponse response, string filePath, string because)
    {
        Assert.False(response.Ok, because);
        Assert.Null(response.Data);
        var error = response.Error!;
        Assert.True(error.Code == "invalid_layout", $"{because}: expected invalid_layout, got {error.Code}: {error.Message}");
        Assert.Contains(filePath, error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(filePath, error.Hint, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("bug in the tool", error.Hint, StringComparison.Ordinal);

        // The edit and preview tools read a private copy; the caller must never be told about it.
        Assert.DoesNotContain("bcwl-edit-", error.Message, StringComparison.Ordinal);
        Assert.DoesNotContain(".bcwl-", error.Message, StringComparison.Ordinal);
    }

    private static T WithFakeConverter<T>(Func<FakePdfConverter, T> body)
    {
        var fake = new FakePdfConverter { Name = "fake", IsAvailable = true, ConversionSucceeds = true };
        LifecycleTools.SelectConverter = _ => fake;
        try
        {
            return body(fake);
        }
        finally
        {
            LifecycleTools.SelectConverter = PdfConverterFactory.Select;
        }
    }

    // ---- not a .docx package at all -> invalid_layout naming the file, explaining .docx vs schema .xml ----

    [Theory]
    [InlineData("empty")]
    [InlineData("text")]
    [InlineData("schema-xml")]
    [InlineData("truncated")]
    public void A_file_that_is_not_a_docx_package_is_invalid_layout_naming_it_in_every_tool_that_opens_one(string kind)
    {
        var dir = PrivateDirectory();
        try
        {
            var bad = WriteNonPackage(dir, kind);
            var badBefore = File.ReadAllBytes(bad);
            var goodLayout = CopyOfCorpusInto(dir, Corpus.SalesInvoice, "good.docx");
            var goodBefore = File.ReadAllBytes(goodLayout);
            var schemaXml = WriteSchemaXml(dir);
            var outputPath = Path.Combine(dir, "created.docx");

            var calls = new (string Tool, Func<ToolResponse> Call)[]
            {
                ("get_layout_info", () => ReadTools.GetLayoutInfo(bad)),
                ("list_dataset_fields", () => ReadTools.ListDatasetFields(bad)),
                ("validate_layout quick", () => ReadTools.ValidateLayout(bad, "quick")),
                ("validate_layout full", () => ReadTools.ValidateLayout(bad, "full")),
                ("insert_field", () => EditTools.InsertField(bad, "/Header/CompanyName", "documentEnd")),
                ("refresh_xml_part layoutPath", () => LifecycleTools.RefreshXmlPart(bad, schemaXml)),
                ("refresh_xml_part newSchemaSource", () => LifecycleTools.RefreshXmlPart(goodLayout, bad)),
                ("create_layout schemaSource", () => LifecycleTools.CreateLayout(bad, outputPath)),
                ("create_layout templatePath", () => LifecycleTools.CreateLayout(schemaXml, outputPath, bad)),
                ("preview_layout", () => WithFakeConverter(fake =>
                {
                    var response = LifecycleTools.PreviewLayout(bad, outputDir: Path.Combine(dir, "preview"));
                    Assert.Equal(0, fake.ConvertCallCount);
                    return response;
                })),
            };

            foreach (var (tool, call) in calls)
            {
                var response = call();
                AssertInvalidLayoutNaming(response, bad, tool);
                Assert.Contains("schema .xml", response.Error!.Hint, StringComparison.Ordinal);
                Assert.Contains("zip container", response.Error!.Hint, StringComparison.Ordinal);
            }

            Assert.Equal(badBefore, File.ReadAllBytes(bad));
            Assert.Equal(goodBefore, File.ReadAllBytes(goodLayout));
            Assert.False(File.Exists(outputPath), "a refused create_layout must write nothing");
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    // ---- a schema .xml that is not well-formed -> invalid_layout naming the file and the line/position ----

    [Theory]
    [InlineData("truncated", "(line 1, position")]
    [InlineData("dtd", "DTD")]
    [InlineData("utf16-declared-on-utf8", "utf-16")]
    [InlineData("docx-bytes", "give it its .docx extension")]
    public void A_schema_xml_that_is_not_well_formed_is_invalid_layout_naming_it(string kind, string hintMustMention)
    {
        var dir = PrivateDirectory();
        try
        {
            var bad = WriteMalformedSchemaXml(dir, kind);
            var goodLayout = CopyOfCorpusInto(dir, Corpus.SalesInvoice, "good.docx");
            var goodBefore = File.ReadAllBytes(goodLayout);
            var outputPath = Path.Combine(dir, "created.docx");

            var calls = new (string Tool, Func<ToolResponse> Call)[]
            {
                ("list_dataset_fields", () => ReadTools.ListDatasetFields(bad)),
                ("create_layout schemaSource", () => LifecycleTools.CreateLayout(bad, outputPath)),
                ("refresh_xml_part newSchemaSource", () => LifecycleTools.RefreshXmlPart(goodLayout, bad)),
            };

            foreach (var (tool, call) in calls)
            {
                var response = call();
                AssertInvalidLayoutNaming(response, bad, tool);
                Assert.Contains("not well-formed XML", response.Error!.Hint, StringComparison.Ordinal);
                Assert.Contains(hintMustMention, response.Error!.Hint, StringComparison.Ordinal);
            }

            Assert.Equal(goodBefore, File.ReadAllBytes(goodLayout));
            Assert.False(File.Exists(outputPath), "a refused create_layout must write nothing");
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public void A_truncated_schema_xml_reports_the_readers_exact_line_and_position()
    {
        var dir = PrivateDirectory();
        try
        {
            var bad = Path.Combine(dir, "multi-line.xml");
            File.WriteAllText(
                bad,
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>\n"
                + "<NavWordReportXmlPart xmlns=\"urn:microsoft-dynamics-nav/reports/Test/1/\">\n"
                + "  <Header>\n"
                + "    <CompanyName>X</CompanyName\n",
                new UTF8Encoding(false));

            var response = ReadTools.ListDatasetFields(bad);

            AssertInvalidLayoutNaming(response, bad, "list_dataset_fields");
            Assert.Matches(@"\(line [45], position \d+\)", response.Error!.Hint);
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    // ---- a .docx whose XML part is damaged (parsed lazily by the SDK) -> invalid_layout ----

    [Fact]
    public void A_layout_with_a_damaged_xml_part_is_invalid_layout_naming_the_layout()
    {
        var dir = PrivateDirectory();
        try
        {
            var damaged = WriteDamagedPackage(dir, "damaged.docx");
            var before = File.ReadAllBytes(damaged);

            var calls = new (string Tool, Func<ToolResponse> Call)[]
            {
                ("get_layout_info", () => ReadTools.GetLayoutInfo(damaged)),
                ("validate_layout quick", () => ReadTools.ValidateLayout(damaged, "quick")),
                ("insert_field", () => EditTools.InsertField(damaged, "/Header/CustomerAddress1", "documentEnd")),
                ("preview_layout", () => WithFakeConverter(_ =>
                    LifecycleTools.PreviewLayout(damaged, outputDir: Path.Combine(dir, "preview")))),
            };

            foreach (var (tool, call) in calls)
            {
                var response = call();
                AssertInvalidLayoutNaming(response, damaged, tool);
                Assert.Contains("damaged", response.Error!.Hint, StringComparison.Ordinal);
                Assert.Contains("(line 1, position", response.Error!.Hint, StringComparison.Ordinal);
            }

            Assert.Equal(before, File.ReadAllBytes(damaged));
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public void CreateLayout_with_a_damaged_templatePath_names_the_template_not_the_schemaSource()
    {
        var dir = PrivateDirectory();
        try
        {
            var template = WriteDamagedPackage(dir, "letterhead.docx");
            var schemaSource = WriteSchemaXml(dir);
            var outputPath = Path.Combine(dir, "created.docx");

            var response = LifecycleTools.CreateLayout(schemaSource, outputPath, template);

            AssertInvalidLayoutNaming(response, template, "create_layout templatePath");
            Assert.DoesNotContain(schemaSource, response.Error!.Hint, StringComparison.OrdinalIgnoreCase);
            Assert.False(File.Exists(outputPath));
            Assert.Empty(Directory.EnumerateFiles(dir, ".bcwl-build-*"));
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    // ---- dataOverridesPath: any unusable file -> invalid_argument for that parameter ----

    [Theory]
    [InlineData("truncated")]
    [InlineData("dtd")]
    [InlineData("wrong-root")]
    [InlineData("wrong-namespace")]
    [InlineData("other-reports-export")]
    public void PreviewLayout_with_an_unusable_dataOverridesPath_is_invalid_argument_listing_the_accepted_shapes(string kind)
    {
        var dir = PrivateDirectory();
        try
        {
            var layout = CopyOfCorpusInto(dir, Corpus.StandardSalesQuote, "quote.docx");
            var overrides = Path.Combine(dir, $"overrides-{kind}.xml");
            var content = kind switch
            {
                "truncated" => "<?xml version=\"1.0\"?><NavWordReportXmlPart xmlns=\"urn:microsoft-dynamics-nav/reports/Standard_Sales_Quote/1304/\"><Header>",
                "dtd" => "<?xml version=\"1.0\"?><!DOCTYPE x [<!ENTITY e \"x\">]><NavWordReportXmlPart xmlns=\"urn:microsoft-dynamics-nav/reports/Standard_Sales_Quote/1304/\"/>",
                "wrong-root" => "<?xml version=\"1.0\"?><SomethingElse/>",
                "wrong-namespace" => "<?xml version=\"1.0\"?><NavWordReportXmlPart xmlns=\"urn:not-bc/\"/>",
                _ => "<?xml version=\"1.0\"?><ReportDataSet name=\"Standard Sales - Invoice\" id=\"1306\"><DataItems/></ReportDataSet>",
            };
            File.WriteAllText(overrides, content, new UTF8Encoding(false));

            var response = WithFakeConverter(fake =>
            {
                var r = LifecycleTools.PreviewLayout(layout, outputDir: Path.Combine(dir, "preview"), dataOverridesPath: overrides);
                Assert.Equal(0, fake.ConvertCallCount);
                return r;
            });

            Assert.False(response.Ok);
            var error = response.Error!;
            Assert.True(error.Code == "invalid_argument", $"expected invalid_argument, got {error.Code}: {error.Message}");
            Assert.Contains(overrides, error.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("dataOverridesPath", error.Hint, StringComparison.Ordinal);
            Assert.Contains("NavWordReportXmlPart", error.Hint, StringComparison.Ordinal);
            Assert.Contains("ReportDataSet", error.Hint, StringComparison.Ordinal);
            Assert.DoesNotContain("bug in the tool", error.Hint, StringComparison.Ordinal);
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    // ---- the internal-bug gate stays internal_error ----

    [Fact]
    public void Guard_keeps_reporting_the_builders_self_check_and_a_stray_FileFormatException_as_internal_error()
    {
        // LayoutBuilder's OpenXmlValidator gate over its own output is unreachable from the public surface
        // (see AttachEverything's remarks), so its mapping is pinned at Guard directly: an internal bug must
        // never be relabelled as the caller's bad file.
        var selfCheck = ToolGuards.Guard(() => throw new BuildSelfCheckFailedException("built layout failed validation"), "x.docx");
        Assert.Equal("internal_error", selfCheck.Error!.Code);

        var stray = ToolGuards.Guard(() => throw new FileFormatException("File contains corrupted data."), "x.docx");
        Assert.Equal("internal_error", stray.Error!.Code);
    }
}
