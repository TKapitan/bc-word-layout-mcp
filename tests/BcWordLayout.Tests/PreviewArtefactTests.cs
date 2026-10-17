using System.Buffers.Binary;
using System.Text.Json;
using BcWordLayout.McpHost;
using BcWordLayout.McpHost.Tools;
using BcWordLayout.Render;
using ModelContextProtocol.Protocol;

namespace BcWordLayout.Tests;

/// <summary>
/// <c>preview_layout</c>'s files on disk are always THIS call's own output (GitHub issue #106). A failed
/// re-conversion used to leave the previous PDF in place, so the folder, and any <c>pdfPath</c> an agent
/// kept, showed the pre-edit render. Every default preview was also named <c>preview.pdf</c>, so
/// <c>render_preview_pages</c> wrote two layouts' pages to one folder under the same
/// <c>preview-page01.png</c>. Every test swaps in <see cref="FakePdfConverter"/>; none reaches Word or
/// LibreOffice.
/// </summary>
/// <remarks>Joins the preview-converter-seam collection: every test here swaps
/// <c>LifecycleTools.SelectConverter</c> (see <see cref="PreviewConverterSeamCollection"/>).</remarks>
[Collection("preview-converter-seam")]
public class PreviewArtefactTests : IDisposable
{
    private readonly string _root = Directory.CreateTempSubdirectory("bcwl-preview-artefacts-").FullName;
    private readonly List<string> _defaultPreviewDirs = new();

    public void Dispose()
    {
        LifecycleTools.SelectConverter = PdfConverterFactory.Select;
        foreach (var dir in _defaultPreviewDirs.Where(Directory.Exists))
        {
            Directory.Delete(dir, recursive: true);
        }

        Directory.Delete(_root, recursive: true);
    }

    /// <summary>A copy of a corpus layout at a fresh path, so its default preview folder is this test's alone.</summary>
    private string LayoutCopy(string subdir, string fileName = "Invoice.docx")
    {
        var dir = Directory.CreateDirectory(Path.Combine(_root, subdir)).FullName;
        var path = Path.Combine(dir, fileName);
        File.Copy(Corpus.Path(Corpus.SalesInvoice), path);
        _defaultPreviewDirs.Add(DefaultPreviewDir(path));
        return path;
    }

    private static string DefaultPreviewDir(string layoutPath) =>
        Path.Combine(Path.GetTempPath(), "bc-word-layout-mcp", LifecycleTools.DefaultPreviewOutputDirName(layoutPath));

    private static PreviewResultDto Preview(string layoutPath, FakePdfConverter converter, string? outputDir = null)
    {
        LifecycleTools.SelectConverter = _ => converter;
        var response = LifecycleTools.PreviewLayout(layoutPath, outputDir: outputDir);
        Assert.True(response.Ok, response.Error?.Message);
        return Assert.IsType<PreviewResultDto>(response.Data);
    }

    private static FakePdfConverter Succeeding(string mediaBox = "0 0 595 842") =>
        new() { Name = "fake-ok", PdfBytes = MinimalPdf.Create(1, mediaBox) };

    private static FakePdfConverter Failing(bool writesPartialOutput = false) =>
        new()
        {
            Name = "fake-timeout",
            ConversionSucceeds = false,
            ConversionErrorMessage = "simulated Word COM timeout",
            WritesPartialOutputOnFailure = writesPartialOutput,
        };

    [Fact]
    public void A_default_preview_names_its_pdf_after_the_layout_hash_and_removes_an_older_versions_preview_pdf()
    {
        var layout = LayoutCopy("named");
        var defaultDir = Directory.CreateDirectory(DefaultPreviewDir(layout)).FullName;
        var legacyPdf = Path.Combine(defaultDir, "preview.pdf");
        File.WriteAllText(legacyPdf, "%PDF-1.4\n%rendered by an older server");

        var dto = Preview(layout, Succeeding());

        Assert.Equal(Path.Combine(defaultDir, $"preview-{LifecycleTools.PreviewPathHash(layout)}.pdf"), dto.PdfPath);
        Assert.True(File.Exists(dto.PdfPath));
        Assert.Equal(Path.Combine(defaultDir, "merged.docx"), dto.MergedDocxPath);
        Assert.False(File.Exists(legacyPdf), "the tool-owned default folder must not keep a stale pre-hash preview.pdf");
    }

    [Fact]
    public void A_failed_re_conversion_in_the_default_folder_leaves_no_pdf_behind()
    {
        // The issue's reproduction: preview, edit, preview again with a converter that fails. The second call
        // returned ok:true / conversionOk:false, but the first render was still on disk at the first pdfPath.
        var layout = LayoutCopy("default-fail");
        var first = Preview(layout, Succeeding());
        Assert.True(File.Exists(first.PdfPath));

        var second = Preview(layout, Failing());

        Assert.False(second.ConversionOk);
        Assert.Equal("simulated Word COM timeout", second.ConversionError);
        Assert.Null(second.PdfPath);
        Assert.False(File.Exists(first.PdfPath), "the previous render must not survive a failed re-conversion");
        Assert.True(File.Exists(second.MergedDocxPath));
    }

    [Fact]
    public void A_failed_re_conversion_into_an_explicit_outputDir_leaves_no_pdf_and_spares_the_callers_own_files()
    {
        var layout = LayoutCopy("explicit-fail");
        var outputDir = Directory.CreateDirectory(Path.Combine(_root, "caller-output")).FullName;
        var callersOwnPdf = Path.Combine(outputDir, "preview.pdf");
        File.WriteAllText(callersOwnPdf, "%PDF-1.4\n%the caller's own file");

        var first = Preview(layout, Succeeding(), outputDir);
        Assert.Equal(Path.Combine(outputDir, $"preview-{LifecycleTools.PreviewPathHash(layout)}.pdf"), first.PdfPath);
        Assert.True(File.Exists(first.PdfPath));

        var second = Preview(layout, Failing(), outputDir);

        Assert.Null(second.PdfPath);
        Assert.False(File.Exists(first.PdfPath));

        // Only this layout's own previous PDF is the tool's to delete in a caller-supplied folder.
        Assert.True(File.Exists(callersOwnPdf));
    }

    [Fact]
    public void Whatever_a_failed_conversion_leaves_at_the_pdf_path_is_removed()
    {
        // A Word export killed on timeout can leave a truncated file at the destination; it is no more
        // this call's render than the previous one was.
        var layout = LayoutCopy("partial");
        var outputDir = Path.Combine(_root, "partial-output");
        var converter = Failing(writesPartialOutput: true);

        var dto = Preview(layout, converter, outputDir);

        Assert.Equal(1, converter.ConvertCallCount);
        Assert.Null(dto.PdfPath);
        Assert.False(File.Exists(converter.LastPdfPath));
    }

    [WindowsOnlyFact]
    public void A_previous_pdf_that_cannot_be_removed_is_reported_and_no_conversion_is_attempted()
    {
        // A PDF viewer holding the previous render open: it cannot be deleted, and the converter could not
        // overwrite it either. The preview says so instead of handing the stale file back as pdfPath.
        var layout = LayoutCopy("locked");
        var first = Preview(layout, Succeeding());
        var second = Succeeding();

        PreviewResultDto dto;
        using (new FileStream(first.PdfPath!, FileMode.Open, FileAccess.Read, FileShare.Read))
        {
            dto = Preview(layout, second);
        }

        Assert.False(dto.ConversionOk);
        Assert.Null(dto.PdfPath);
        Assert.Equal("fake-ok", dto.ConverterUsed);
        Assert.Contains("could not be removed", dto.ConversionError, StringComparison.Ordinal);
        Assert.Contains(first.PdfPath!, dto.ConversionError, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(0, second.ConvertCallCount);
    }

    [Fact]
    public void Default_previews_of_two_layouts_render_their_pages_into_one_folder_under_distinct_names()
    {
        // Two layouts that even share a file name, each previewed into its own default folder, then both
        // rendered into ONE outputDir. With both PDFs named preview.pdf the second render overwrote the
        // first's preview-page01.png, and the first layout's returned path showed the second layout's page.
        // Different page sizes make the two renders tell apart by their PNG dimensions.
        var layoutA = LayoutCopy("appA");
        var layoutB = LayoutCopy("appB");
        var pdfA = Preview(layoutA, Succeeding("0 0 595 842")).PdfPath!;
        var pdfB = Preview(layoutB, Succeeding("0 0 842 595")).PdfPath!;
        var pngDir = Path.Combine(_root, "pngs");

        var pageA = RenderFirstPage(pdfA, pngDir);
        var pageB = RenderFirstPage(pdfB, pngDir);

        Assert.NotEqual(pageA.Path, pageB.Path, StringComparer.OrdinalIgnoreCase);
        Assert.Equal(2, Directory.GetFiles(pngDir).Length);

        // Layout A's PNG on disk is still layout A's page.
        Assert.NotEqual(pageA.WidthPx, pageB.WidthPx);
        Assert.Equal(pageA.WidthPx, PngWidth(pageA.Path));
        Assert.Equal(pageB.WidthPx, PngWidth(pageB.Path));
    }

    private static (string Path, int WidthPx) RenderFirstPage(string pdfPath, string outputDir)
    {
        var result = LifecycleTools.RenderPreviewPages(pdfPath, maxPages: 1, dpi: 36, outputDir: outputDir, inlineImages: false);
        var envelope = Assert.IsType<TextContentBlock>(Assert.Single(result.Content));
        using var json = JsonDocument.Parse(envelope.Text);
        Assert.True(json.RootElement.GetProperty("ok").GetBoolean(), envelope.Text);
        var page = json.RootElement.GetProperty("data").GetProperty("pages")[0];
        return (page.GetProperty("path").GetString()!, page.GetProperty("widthPx").GetInt32());
    }

    /// <summary>A PNG's width, read from its IHDR chunk (big-endian, right after the 8-byte signature and the chunk header).</summary>
    private static int PngWidth(string pngPath) =>
        BinaryPrimitives.ReadInt32BigEndian(File.ReadAllBytes(pngPath).AsSpan(16, 4));
}
