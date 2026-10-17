using BcWordLayout.Domain;
using BcWordLayout.McpHost;
using BcWordLayout.McpHost.Tools;

namespace BcWordLayout.Tests;

/// <summary>
/// GitHub issue #89: <c>create_layout</c> keeps only <c>schemaSource</c>'s dataset part, so an
/// <c>outputPath</c> that names <c>schemaSource</c> itself used to replace an authored layout with a blank one
/// while returning <c>ok:true</c>. It is now refused (<c>invalid_argument</c>, file byte-identical) whenever that
/// file carries bound controls, however the path is spelled; the supported aliasing, scaffolding a
/// compiler-generated layout that has the dataset part but no controls yet, keeps working. Aliasing
/// <c>templatePath</c> is unaffected (the template's body IS the output's body, so nothing authored is
/// dropped; <c>LayoutBuilderTests</c> pins that case and its <c>template_not_unbound</c> refusal). Every test
/// works on copies in its own private directory.
/// </summary>
public class CreateLayoutAliasingTests
{
    private static string PrivateDirectory() => Directory.CreateTempSubdirectory("bcwl-alias-").FullName;

    private static string CopyOfCorpusInto(string dir, string corpusFile, string fileName)
    {
        var path = Path.Combine(dir, fileName);
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    private static int BoundControlCount(string layoutPath) =>
        LayoutReader.Read(layoutPath).Controls.Count(c => c.XPath is not null);

    private static void AssertRefusedAsAliasing(ToolResponse response, string layoutPath, int boundControls)
    {
        Assert.False(response.Ok);
        Assert.Null(response.Data);
        var error = response.Error!;
        Assert.Equal("invalid_argument", error.Code);
        Assert.Contains(Path.GetFullPath(layoutPath), error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("same file as schemaSource", error.Message, StringComparison.Ordinal);
        Assert.Contains($"{boundControls} bound content control(s)", error.Message, StringComparison.Ordinal);
        Assert.Contains("Nothing was written", error.Message, StringComparison.Ordinal);
        Assert.Contains("refresh_xml_part", error.Hint, StringComparison.Ordinal);
        Assert.Contains("different outputPath", error.Hint, StringComparison.Ordinal);
    }

    [Fact]
    public void CreateLayout_over_its_own_authored_schemaSource_is_refused_and_leaves_the_layout_intact()
    {
        var dir = PrivateDirectory();
        try
        {
            var quote = CopyOfCorpusInto(dir, Corpus.StandardSalesQuote, "quote.docx");
            var before = File.ReadAllBytes(quote);
            var boundControls = BoundControlCount(quote);
            Assert.True(boundControls > 0, "fixture must be an authored layout");

            var response = LifecycleTools.CreateLayout(quote, quote);

            AssertRefusedAsAliasing(response, quote, boundControls);
            Assert.Equal(before, File.ReadAllBytes(quote));
            Assert.Equal(boundControls, BoundControlCount(quote));
            Assert.Empty(Directory.EnumerateFiles(dir, ".bcwl-*"));
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public void Every_spelling_of_the_same_path_is_recognized_as_schemaSource()
    {
        var dir = PrivateDirectory();
        try
        {
            var quote = CopyOfCorpusInto(dir, Corpus.StandardSalesQuote, "quote.docx");
            var before = File.ReadAllBytes(quote);
            var boundControls = BoundControlCount(quote);

            var spellings = new[]
            {
                Path.Combine(dir, ".", "quote.docx"),
                Path.Combine(dir, "sub", "..", "quote.docx"),
                quote.Replace('\\', '/'),
            }.ToList();
            if (OperatingSystem.IsWindows())
            {
                spellings.Add(quote.ToUpperInvariant());
            }

            foreach (var outputPath in spellings)
            {
                var response = LifecycleTools.CreateLayout(quote, outputPath);
                AssertRefusedAsAliasing(response, quote, boundControls);
            }

            Assert.Equal(before, File.ReadAllBytes(quote));
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    [WindowsOnlyFact]
    public void A_junction_or_8_3_spelling_of_schemaSource_is_recognized_too()
    {
        // GitHub issue #92 (noted on #89): Path.GetFullPath does not follow a junction, so before the same-file
        // check used PathIdentity that spelling overwrote the authored layout. The 8.3 spelling was already
        // expanded by GetFullPath and stays recognized.
        var dir = PrivateDirectory();
        var real = Directory.CreateDirectory(Path.Combine(dir, "real-folder-with-a-long-name")).FullName;
        var link = Path.Combine(dir, "link");
        try
        {
            var quote = CopyOfCorpusInto(real, Corpus.StandardSalesQuote, "quote.docx");
            var before = File.ReadAllBytes(quote);
            var boundControls = BoundControlCount(quote);
            FileSystemLinks.CreateJunction(link, real);

            var spellings = new List<string> { Path.Combine(link, "quote.docx") };
            if (FileSystemLinks.ShortPathOrNull(quote) is { } shortSpelling)
            {
                spellings.Add(shortSpelling);
            }

            foreach (var outputPath in spellings)
            {
                var response = LifecycleTools.CreateLayout(quote, outputPath);

                Assert.False(response.Ok);
                Assert.Equal("invalid_argument", response.Error!.Code);
                Assert.Contains("same file as schemaSource", response.Error.Message, StringComparison.Ordinal);
                Assert.Contains($"{boundControls} bound content control(s)", response.Error.Message, StringComparison.Ordinal);
            }

            Assert.Equal(before, File.ReadAllBytes(quote));
            Assert.Empty(Directory.EnumerateFiles(real, ".bcwl-*"));
        }
        finally
        {
            FileSystemLinks.DeleteJunction(link);
            Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public void A_different_templatePath_does_not_make_overwriting_an_authored_schemaSource_acceptable()
    {
        // The output's body would be the template's, so the authored body is still discarded.
        var dir = PrivateDirectory();
        try
        {
            var quote = CopyOfCorpusInto(dir, Corpus.StandardSalesQuote, "quote.docx");
            var before = File.ReadAllBytes(quote);
            var shell = Path.Combine(dir, "shell.docx");
            LayoutBuilder.Create(Corpus.Path(Corpus.StandardStatement), shell);

            var response = LifecycleTools.CreateLayout(quote, quote, shell);

            AssertRefusedAsAliasing(response, quote, BoundControlCount(quote));
            Assert.Equal(before, File.ReadAllBytes(quote));
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public void LayoutBuilder_reports_the_refusal_as_an_ArgumentException_for_outputPath()
    {
        var dir = PrivateDirectory();
        try
        {
            var quote = CopyOfCorpusInto(dir, Corpus.StandardSalesQuote, "quote.docx");

            var ex = Assert.Throws<ArgumentException>(() => LayoutBuilder.Create(quote, quote));

            Assert.Equal("outputPath", ex.ParamName);
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public void Scaffolding_a_schemaSource_with_no_bound_controls_in_place_still_works()
    {
        // The supported aliasing: a compiler-generated layout carries the dataset part and no controls yet,
        // and create_layout over it in place adds the scaffolding (header/footer, styles, settings).
        var dir = PrivateDirectory();
        try
        {
            var generated = Path.Combine(dir, "generated.docx");
            var first = LayoutBuilder.Create(Corpus.Path(Corpus.StandardSalesQuote), generated);
            Assert.Equal(0, BoundControlCount(generated));

            var response = LifecycleTools.CreateLayout(generated, generated);

            Assert.True(response.Ok, response.Error?.Message);
            var dto = Assert.IsType<CreateResultDto>(response.Data);
            Assert.Equal(Path.GetFullPath(generated), dto.OutputPath);
            Assert.Equal(first.ReportName, dto.ReportName);
            Assert.Equal(first.Namespace, dto.Namespace);
            Assert.True(dto.QuickValidation.Passed);
            Assert.Equal(first.ReportName, SchemaProvider.FromLayout(generated).Report.ReportName);
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public void When_templatePath_is_the_same_authored_file_too_the_template_rules_still_decide()
    {
        // schemaSource == templatePath == outputPath keeps the file's own body, so nothing authored would be
        // dropped by the build itself; what breaks is its controls' bindings, which template_not_unbound
        // already refuses (with its own refresh_xml_part guidance).
        var dir = PrivateDirectory();
        try
        {
            var quote = CopyOfCorpusInto(dir, Corpus.StandardSalesQuote, "quote.docx");
            var before = File.ReadAllBytes(quote);

            var response = LifecycleTools.CreateLayout(quote, quote, quote);

            Assert.False(response.Ok);
            Assert.Equal("template_not_unbound", response.Error!.Code);
            Assert.Equal(before, File.ReadAllBytes(quote));
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }
}
