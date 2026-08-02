using PDFtoImage;

namespace BcWordLayout.Render;

/// <summary>Options controlling a single <see cref="PdfRasterizer.Rasterize"/> call.</summary>
public sealed class PdfRasterizeOptions
{
    /// <summary>
    /// Hard ceiling on how many pages one call may render, regardless of what the caller asks for —
    /// mirrors the row-cap philosophy in the merge layer: an oversized request is clamped and REPORTED
    /// (<see cref="PdfRasterizeResult.Truncated"/>), never silently honored. Ten pages at
    /// <see cref="MaxDpi"/>-bounded resolution keeps the worst-case response payload bounded; a preview
    /// PDF an agent needs to LOOK at is nearly always its first page or two.
    /// </summary>
    public const int MaxPageCap = 10;

    /// <summary>Inclusive DPI bounds. 36 is still legible as a thumbnail; 300 is print resolution.</summary>
    public const int MinDpi = 36;

    /// <summary>Inclusive upper DPI bound — see <see cref="MinDpi"/>.</summary>
    public const int MaxDpi = 300;

    /// <summary>
    /// Refuse input PDFs larger than this (50 MB). A preview PDF produced by this codebase's own
    /// converters is orders of magnitude smaller; anything bigger is a wrong-file mistake, not a preview.
    /// </summary>
    public const long MaxPdfBytes = 50 * 1024 * 1024;

    /// <summary>1-based page to start rendering from. Default 1 (the first page).</summary>
    public int FirstPage { get; init; } = 1;

    /// <summary>How many pages to render, starting at <see cref="FirstPage"/>. Clamped to <see cref="MaxPageCap"/>.</summary>
    public int MaxPages { get; init; } = 3;

    /// <summary>
    /// Render resolution. Default 120: an A4 page comes out ~992×1403 px — comfortably readable for a
    /// vision model without wasting payload on print-resolution detail. Clamped to
    /// [<see cref="MinDpi"/>, <see cref="MaxDpi"/>].
    /// </summary>
    public int Dpi { get; init; } = 120;
}

/// <summary>One rendered page: PNG bytes plus the pixel dimensions they decode to.</summary>
public sealed record PdfRasterizedPage(int PageNumber, byte[] PngBytes, int WidthPx, int HeightPx);

/// <summary>
/// The outcome of one PDF → PNG-pages rasterization attempt. Mirrors <see cref="PdfConversionResult"/>'s
/// contract: <see cref="PdfRasterizer.Rasterize"/> never throws, it always returns one of these.
/// </summary>
public sealed class PdfRasterizeResult
{
    /// <summary>Whether at least the requested (clamped) pages rendered successfully.</summary>
    public required bool Ok { get; init; }

    /// <summary>Total pages in the source PDF. Set only when the PDF could be opened at all.</summary>
    public int? PageCount { get; init; }

    /// <summary>The rendered pages, in ascending page order. Empty when <see cref="Ok"/> is false.</summary>
    public IReadOnlyList<PdfRasterizedPage> Pages { get; init; } = [];

    /// <summary>
    /// True when the document has pages beyond the last one rendered (whether the request or the
    /// per-call cap stopped short). Callers surface this so an agent never mistakes "the first N pages"
    /// for "the whole document". Receiving everything from <c>FirstPage</c> to the document's end is a
    /// complete answer and is NOT truncation, even if more pages were requested than existed.
    /// </summary>
    public bool Truncated { get; init; }

    /// <summary>The effective DPI the pages were rendered at (after clamping).</summary>
    public int EffectiveDpi { get; init; }

    /// <summary>Human-actionable description of what went wrong. Set only when <see cref="Ok"/> is false.</summary>
    public string? Error { get; init; }
}

/// <summary>
/// Renders the leading pages of a PDF (typically one produced by an <see cref="IPdfConverter"/>) to PNG
/// images via PDFium (the PDFtoImage package). Unlike the converters this has no external-install
/// dependency — PDFium ships as a native library inside the package — so there is no
/// <c>IsAvailable</c> notion and no factory: it either renders or reports a structured failure.
/// </summary>
public static class PdfRasterizer
{
    /// <summary>
    /// Rasterizes up to <see cref="PdfRasterizeOptions.MaxPages"/> pages of <paramref name="pdfPath"/>
    /// starting at <see cref="PdfRasterizeOptions.FirstPage"/>. Never throws: missing/oversized/non-PDF
    /// input and PDFium render failures all come back as <see cref="PdfRasterizeResult.Ok"/> = false with
    /// a human-actionable <see cref="PdfRasterizeResult.Error"/>.
    /// </summary>
    public static PdfRasterizeResult Rasterize(string pdfPath, PdfRasterizeOptions? options = null)
    {
        options ??= new PdfRasterizeOptions();

        if (!File.Exists(pdfPath))
        {
            return Fail($"pdfPath does not point to an existing file: '{pdfPath}'.");
        }

        if (!PdfFileValidation.LooksLikePdf(pdfPath))
        {
            return Fail($"'{pdfPath}' is not a PDF (missing %PDF header) or could not be read.");
        }

        byte[] pdfBytes;
        try
        {
            var length = new FileInfo(pdfPath).Length;
            if (length > PdfRasterizeOptions.MaxPdfBytes)
            {
                return Fail(
                    $"'{pdfPath}' is {length / (1024 * 1024)} MB, above the {PdfRasterizeOptions.MaxPdfBytes / (1024 * 1024)} MB "
                    + "rasterization limit for preview PDFs.");
            }

            pdfBytes = File.ReadAllBytes(pdfPath);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return Fail($"Could not read '{pdfPath}': {ex.Message}");
        }

        var firstPage = Math.Max(1, options.FirstPage);
        var requestedPages = Math.Max(1, options.MaxPages);
        var cappedPages = Math.Min(requestedPages, PdfRasterizeOptions.MaxPageCap);
        var dpi = Math.Clamp(options.Dpi, PdfRasterizeOptions.MinDpi, PdfRasterizeOptions.MaxDpi);

        // CA1416: PDFtoImage annotates its API as supported on Windows/Linux/macOS/mobile — every desktop
        // platform this server can run on at all (and in practice it is Windows-only, per the converters).
        // The analyzer flags the call anyway because the supported list is not literally "all platforms";
        // there is no reachable unsupported platform to guard against here.
#pragma warning disable CA1416
        try
        {
            var pageCount = Conversion.GetPageCount(pdfBytes);
            if (firstPage > pageCount)
            {
                return new PdfRasterizeResult
                {
                    Ok = false,
                    PageCount = pageCount,
                    EffectiveDpi = dpi,
                    Error = $"firstPage {firstPage} is beyond the document's last page ({pageCount}).",
                };
            }

            var lastPage = Math.Min(pageCount, firstPage + cappedPages - 1);
            var renderOptions = new RenderOptions { Dpi = dpi };
            var pages = new List<PdfRasterizedPage>(lastPage - firstPage + 1);
            for (var page = firstPage; page <= lastPage; page++)
            {
                using var pngStream = new MemoryStream();
                // PDFtoImage's Index parameter is 0-based.
                Conversion.SavePng(pngStream, pdfBytes, page - 1, options: renderOptions);
                var pngBytes = pngStream.ToArray();
                var (widthPx, heightPx) = ReadPngDimensions(pngBytes);
                pages.Add(new PdfRasterizedPage(page, pngBytes, widthPx, heightPx));
            }

            return new PdfRasterizeResult
            {
                Ok = true,
                PageCount = pageCount,
                Pages = pages,
                // "Truncated" means exactly one thing: the document has pages BEYOND the last one
                // rendered. Asking for more pages than exist and receiving the whole remainder is a
                // complete answer, not a truncated one — over-clamping only matters when it actually
                // hid a page.
                Truncated = lastPage < pageCount,
                EffectiveDpi = dpi,
            };
        }
        catch (Exception ex)
        {
            // PDFium reports corrupt/encrypted/unsupported PDFs as exceptions from the wrapper; there is
            // no stable public exception taxonomy to catch more narrowly, and the converters' own
            // "never throw" contract is worth more here than rethrowing something no caller can act on.
            return Fail($"PDF rasterization failed for '{pdfPath}': {ex.Message}");
        }
#pragma warning restore CA1416

        static PdfRasterizeResult Fail(string error) => new() { Ok = false, Error = error };
    }

    /// <summary>
    /// Reads width/height straight from the fixed-offset IHDR chunk every valid PNG starts with —
    /// SavePng just produced these bytes, so this is a header read, not a decode.
    /// </summary>
    private static (int Width, int Height) ReadPngDimensions(byte[] png)
    {
        // 8-byte signature + 4-byte length + "IHDR" = width at offset 16, height at 20 (big-endian).
        if (png.Length < 24)
        {
            return (0, 0);
        }

        var width = (png[16] << 24) | (png[17] << 16) | (png[18] << 8) | png[19];
        var height = (png[20] << 24) | (png[21] << 16) | (png[22] << 8) | png[23];
        return (width, height);
    }
}
