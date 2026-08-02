namespace BcWordLayout.Render;

/// <summary>Which <see cref="IPdfConverter"/> a caller wants <see cref="PdfConverterFactory.Select"/> to resolve.</summary>
public enum PdfConverterKind
{
    /// <summary>Prefer Word COM when available; else LibreOffice; else a converter that fails cleanly.</summary>
    Auto,

    /// <summary>Word COM automation only (Windows, requires a local Microsoft Word install).</summary>
    Word,

    /// <summary>LibreOffice headless only (cross-platform, requires a local LibreOffice install).</summary>
    LibreOffice,
}

/// <summary>Options controlling a single <see cref="IPdfConverter.Convert"/> call.</summary>
public sealed class PdfConversionOptions
{
    /// <summary>
    /// Wall-clock budget for the whole conversion. If exceeded, the converter kills its worker process
    /// (for Word COM, the specific tracked WINWORD.EXE instance) and returns a timeout failure rather than
    /// hanging forever.
    /// </summary>
    public TimeSpan Timeout { get; init; } = TimeSpan.FromSeconds(120);

    /// <summary>Whether a file already at the destination path may be overwritten.</summary>
    public bool Overwrite { get; init; } = true;
}

/// <summary>
/// The outcome of one .docx → PDF conversion attempt. A conforming <see cref="IPdfConverter"/> never
/// throws; it always returns one of these, success or failure.
/// </summary>
public sealed class PdfConversionResult
{
    /// <summary>Whether the PDF was produced and passed the output sanity checks (exists, starts with <c>%PDF</c>).</summary>
    public required bool Ok { get; init; }

    /// <summary>Absolute path to the produced PDF. Set only when <see cref="Ok"/> is true.</summary>
    public string? PdfPath { get; init; }

    /// <summary>Which converter ran, e.g. <c>"word-com"</c> or <c>"libreoffice"</c> — see <see cref="IPdfConverter.Name"/>.</summary>
    public required string Converter { get; init; }

    /// <summary>Human-actionable description of what went wrong. Set only when <see cref="Ok"/> is false.</summary>
    public string? Error { get; init; }

    /// <summary>Wall-clock time the conversion attempt took.</summary>
    public TimeSpan Duration { get; init; }

    /// <summary>Builds a successful result.</summary>
    public static PdfConversionResult Success(string converter, string pdfPath, TimeSpan duration) =>
        new() { Ok = true, Converter = converter, PdfPath = pdfPath, Duration = duration };

    /// <summary>Builds a failed result. <paramref name="error"/> should be specific enough to act on.</summary>
    public static PdfConversionResult Failure(string converter, string error, TimeSpan duration = default) =>
        new() { Ok = false, Converter = converter, Error = error, Duration = duration };
}

/// <summary>A backend capable of converting a merged BC Word layout preview (.docx) to PDF.</summary>
public interface IPdfConverter
{
    /// <summary>Short machine-readable identity, e.g. <c>"word-com"</c> or <c>"libreoffice"</c>.</summary>
    string Name { get; }

    /// <summary>
    /// Whether this converter's dependency (a Word install / the <c>soffice</c> executable) is present on
    /// this machine. Reflects live machine state — evaluate it again after installing/removing a
    /// dependency rather than caching the result.
    /// </summary>
    bool IsAvailable { get; }

    /// <summary>
    /// Converts <paramref name="docxPath"/> to a PDF at <paramref name="pdfPath"/>. Never throws: every
    /// failure mode (missing input, missing dependency, timeout, converter crash, invalid output) is
    /// reported as a structured <see cref="PdfConversionResult"/> with <c>Ok = false</c> and a
    /// human-actionable <see cref="PdfConversionResult.Error"/>.
    /// </summary>
    PdfConversionResult Convert(string docxPath, string pdfPath, PdfConversionOptions? options = null);
}
