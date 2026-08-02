using System.Text;
using BcWordLayout.McpHost;
using BcWordLayout.McpHost.Tools;
using BcWordLayout.Render;
using ModelContextProtocol.Protocol;

namespace BcWordLayout.Tests;

/// <summary>
/// Builds a minimal but structurally valid PDF (correct xref offsets, ASCII-only so string length ==
/// byte offset) with N empty A4 pages — enough for PDFium to open and render it for real, without
/// needing Word/LibreOffice on the test machine or a binary fixture in the repo.
/// </summary>
internal static class MinimalPdf
{
    internal static byte[] Create(int pageCount)
    {
        var objects = new List<string>
        {
            "<< /Type /Catalog /Pages 2 0 R >>",
            $"<< /Type /Pages /Kids [{string.Join(" ", Enumerable.Range(0, pageCount).Select(i => $"{3 + i} 0 R"))}] /Count {pageCount} >>",
        };
        objects.AddRange(Enumerable.Range(0, pageCount)
            .Select(_ => "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] >>"));

        var body = new StringBuilder("%PDF-1.4\n");
        var offsets = new List<int>();
        for (var i = 0; i < objects.Count; i++)
        {
            offsets.Add(body.Length);
            body.Append($"{i + 1} 0 obj\n{objects[i]}\nendobj\n");
        }

        var xrefStart = body.Length;
        body.Append($"xref\n0 {objects.Count + 1}\n0000000000 65535 f \n");
        foreach (var offset in offsets)
        {
            body.Append($"{offset:D10} 00000 n \n");
        }

        body.Append($"trailer\n<< /Size {objects.Count + 1} /Root 1 0 R >>\nstartxref\n{xrefStart}\n%%EOF");
        return Encoding.ASCII.GetBytes(body.ToString());
    }

    internal static string CreateFile(string directory, int pageCount)
    {
        var path = Path.Combine(directory, $"minimal-{pageCount}p-{Guid.NewGuid():N}.pdf");
        File.WriteAllBytes(path, Create(pageCount));
        return path;
    }
}

/// <summary>
/// Tests for <see cref="PdfRasterizer"/> against the REAL PDFium native library (no seam/fake): the
/// package ships its own native binary, so unlike the converters there is no machine-dependent install
/// to guard against — these always run.
/// </summary>
public class PdfRasterizerTests : IDisposable
{
    private static readonly byte[] PngMagic = [0x89, 0x50, 0x4E, 0x47];

    private readonly string _dir = Directory.CreateTempSubdirectory("bcwl-rasterizer-tests-").FullName;

    public void Dispose() => Directory.Delete(_dir, recursive: true);

    [Fact]
    public void Rasterize_missing_file_fails_cleanly()
    {
        var result = PdfRasterizer.Rasterize(Path.Combine(_dir, "missing.pdf"));

        Assert.False(result.Ok);
        Assert.Contains("existing file", result.Error);
        Assert.Empty(result.Pages);
    }

    [Fact]
    public void Rasterize_non_pdf_input_fails_cleanly()
    {
        var path = Path.Combine(_dir, "not-a-pdf.pdf");
        File.WriteAllText(path, "just text, no PDF header");

        var result = PdfRasterizer.Rasterize(path);

        Assert.False(result.Ok);
        Assert.Contains("%PDF", result.Error);
    }

    [Fact]
    public void Rasterize_renders_real_png_pages_with_dimensions()
    {
        var path = MinimalPdf.CreateFile(_dir, 2);

        var result = PdfRasterizer.Rasterize(path, new PdfRasterizeOptions { Dpi = 72 });

        Assert.True(result.Ok, result.Error);
        Assert.Equal(2, result.PageCount);
        Assert.Equal(2, result.Pages.Count);
        Assert.False(result.Truncated);
        Assert.Equal(72, result.EffectiveDpi);
        foreach (var page in result.Pages)
        {
            Assert.Equal(PngMagic, page.PngBytes.Take(4));
            // A4 at 72 dpi is exactly the PDF's 595×842 point MediaBox.
            Assert.Equal(595, page.WidthPx);
            Assert.Equal(842, page.HeightPx);
        }

        Assert.Equal([1, 2], result.Pages.Select(p => p.PageNumber));
    }

    [Fact]
    public void Rasterize_reports_truncation_when_document_has_more_pages_than_requested()
    {
        var path = MinimalPdf.CreateFile(_dir, 3);

        var result = PdfRasterizer.Rasterize(path, new PdfRasterizeOptions { MaxPages = 2, Dpi = 36 });

        Assert.True(result.Ok, result.Error);
        Assert.Equal(3, result.PageCount);
        Assert.Equal(2, result.Pages.Count);
        Assert.True(result.Truncated);
    }

    [Fact]
    public void Rasterize_clamps_an_oversized_page_request_to_the_hard_cap()
    {
        var path = MinimalPdf.CreateFile(_dir, 12);

        var result = PdfRasterizer.Rasterize(path, new PdfRasterizeOptions { MaxPages = 50, Dpi = 36 });

        Assert.True(result.Ok, result.Error);
        Assert.Equal(PdfRasterizeOptions.MaxPageCap, result.Pages.Count);
        Assert.True(result.Truncated);
    }

    [Fact]
    public void Rasterize_pages_from_firstPage_onward()
    {
        var path = MinimalPdf.CreateFile(_dir, 4);

        var result = PdfRasterizer.Rasterize(path, new PdfRasterizeOptions { FirstPage = 3, MaxPages = 5, Dpi = 36 });

        Assert.True(result.Ok, result.Error);
        Assert.Equal([3, 4], result.Pages.Select(p => p.PageNumber));
        // Requested 5 pages from page 3 and got everything through the document's end — a COMPLETE
        // answer (pages 1-2 were excluded by the caller's own firstPage, not cut by the tool).
        Assert.False(result.Truncated);
    }

    [Fact]
    public void Rasterize_firstPage_beyond_the_document_fails_with_the_page_count()
    {
        var path = MinimalPdf.CreateFile(_dir, 2);

        var result = PdfRasterizer.Rasterize(path, new PdfRasterizeOptions { FirstPage = 5 });

        Assert.False(result.Ok);
        Assert.Equal(2, result.PageCount);
        Assert.Contains("last page", result.Error);
    }

    [Fact]
    public void Rasterize_clamps_dpi_into_bounds()
    {
        var path = MinimalPdf.CreateFile(_dir, 1);

        var result = PdfRasterizer.Rasterize(path, new PdfRasterizeOptions { Dpi = 9999, MaxPages = 1 });

        Assert.True(result.Ok, result.Error);
        Assert.Equal(PdfRasterizeOptions.MaxDpi, result.EffectiveDpi);
    }
}

/// <summary>
/// Tests for the <c>render_preview_pages</c> MCP tool surface: the one tool returning a raw
/// <see cref="CallToolResult"/> (JSON envelope text block + one image block per page) instead of a
/// serialized <see cref="ToolResponse"/>. No preview-converter seam is touched (nothing here calls
/// <c>PreviewLayout</c>), so this class needs no seam collection membership.
/// </summary>
public class RenderPreviewPagesToolTests : IDisposable
{
    private readonly string _dir = Directory.CreateTempSubdirectory("bcwl-render-pages-tests-").FullName;

    public void Dispose() => Directory.Delete(_dir, recursive: true);

    [Fact]
    public void Missing_pdf_returns_a_single_text_block_with_the_standard_failure_envelope()
    {
        var result = LifecycleTools.RenderPreviewPages(Path.Combine(_dir, "gone.pdf"));

        var block = Assert.Single(result.Content);
        var text = Assert.IsType<TextContentBlock>(block);
        Assert.Contains("\"ok\":false", text.Text);
        Assert.Contains("file_not_found", text.Text);
        Assert.Contains("preview_layout", text.Text); // the recovery hint names the tool to re-run
    }

    [Fact]
    public void Unrenderable_pdf_returns_a_rasterize_failed_envelope()
    {
        var path = Path.Combine(_dir, "bad.pdf");
        File.WriteAllText(path, "%PDF-1.4 but nothing else that a PDF needs");

        var result = LifecycleTools.RenderPreviewPages(path);

        var block = Assert.Single(result.Content);
        var text = Assert.IsType<TextContentBlock>(block);
        Assert.Contains("\"ok\":false", text.Text);
        Assert.Contains("rasterize_failed", text.Text);
    }

    [Fact]
    public void Success_returns_envelope_text_block_then_one_png_image_block_per_page()
    {
        var path = MinimalPdf.CreateFile(_dir, 3);

        var result = LifecycleTools.RenderPreviewPages(path, maxPages: 2, dpi: 72);

        Assert.Equal(3, result.Content.Count);

        var envelope = Assert.IsType<TextContentBlock>(result.Content[0]);
        Assert.Contains("\"ok\":true", envelope.Text);
        Assert.Contains("\"pageCount\":3", envelope.Text);
        Assert.Contains("\"pagesRendered\":2", envelope.Text);
        Assert.Contains("\"truncated\":true", envelope.Text);

        foreach (var block in result.Content.Skip(1))
        {
            var image = Assert.IsType<ImageContentBlock>(block);
            Assert.Equal("image/png", image.MimeType);
            // Data (the wire field) is base64 UTF-8; DecodedData round-trips it back to the raw PNG.
            Assert.True(image.Data.Length > 0);
            Assert.Equal([(byte)0x89, (byte)'P', (byte)'N', (byte)'G'], image.DecodedData.ToArray().Take(4));
        }
    }

    [Fact]
    public void Envelope_JSON_uses_the_same_serialization_shape_as_every_other_tools_envelope()
    {
        var path = MinimalPdf.CreateFile(_dir, 1);

        var result = LifecycleTools.RenderPreviewPages(path, maxPages: 1, dpi: 36);

        var envelope = Assert.IsType<TextContentBlock>(result.Content[0]);
        // Same camelCase envelope keys agents already parse from the other tools ({ok, data, error}) —
        // this pins the SerializeEnvelope helper to McpJsonUtilities.DefaultOptions-compatible output.
        Assert.StartsWith("{\"ok\":true,\"data\":{", envelope.Text);
        Assert.Contains("\"pdfPath\":", envelope.Text);
        Assert.Contains("\"effectiveDpi\":36", envelope.Text);
    }
}
