using BcWordLayout.Render;

namespace BcWordLayout.Tests;

/// <summary>
/// A fake <see cref="IPdfConverter"/> for deterministic <c>preview_layout</c> tests, swapped in via <see cref="BcWordLayout.McpHost.Tools.LifecycleTools.SelectConverter"/> so the
/// tool's PDF-conversion outcome handling can be asserted without depending on whether Word/LibreOffice is
/// actually installed on the machine running the suite. Every constructed instance is configured up front
/// (<see cref="IsAvailable"/>/<see cref="ConversionSucceeds"/>/<see cref="ConversionErrorMessage"/>) and then
/// records what <c>preview_layout</c> actually did with it (<see cref="ConvertCallCount"/>/
/// <see cref="LastDocxPath"/>/<see cref="LastPdfPath"/>), so a test can assert both the tool's reported
/// outcome AND that the converter was genuinely invoked with the real merged working copy rather than
/// asserting the outcome shape alone.
/// </summary>
internal sealed class FakePdfConverter : IPdfConverter
{
    /// <summary>Reported as <c>converterUsed</c> in <see cref="BcWordLayout.McpHost.PreviewResultDto"/>.</summary>
    public string Name { get; init; } = "fake";

    /// <summary>Reported as <c>converterAvailable</c>. Independent of whether <see cref="Convert"/> actually
    /// succeeds - mirrors the real converters, whose <c>Convert</c> is called unconditionally by
    /// <c>preview_layout</c> regardless of <see cref="IsAvailable"/> (each real converter fails cleanly from
    /// inside <c>Convert</c> itself when its own dependency is missing, rather than the tool branching on
    /// <see cref="IsAvailable"/> first) - see <c>LifecycleTools.PreviewLayout</c>.</summary>
    public bool IsAvailable { get; init; } = true;

    /// <summary>Whether <see cref="Convert"/> reports success (and writes a fake PDF) or a clean failure.</summary>
    public bool ConversionSucceeds { get; init; } = true;

    /// <summary>Error text <see cref="Convert"/> reports when <see cref="ConversionSucceeds"/> is false.</summary>
    public string ConversionErrorMessage { get; init; } = "simulated conversion failure";

    /// <summary>How many times <see cref="Convert"/> was actually called.</summary>
    public int ConvertCallCount { get; private set; }

    /// <summary>The <c>docxPath</c> argument of the most recent <see cref="Convert"/> call, if any.</summary>
    public string? LastDocxPath { get; private set; }

    /// <summary>The <c>pdfPath</c> argument of the most recent <see cref="Convert"/> call, if any.</summary>
    public string? LastPdfPath { get; private set; }

    public PdfConversionResult Convert(string docxPath, string pdfPath, PdfConversionOptions? options = null)
    {
        ConvertCallCount++;
        LastDocxPath = docxPath;
        LastPdfPath = pdfPath;

        if (!ConversionSucceeds)
        {
            return PdfConversionResult.Failure(Name, ConversionErrorMessage);
        }

        // Write a minimal-but-recognizable fake PDF so "the PDF file exists / starts with %PDF" assertions
        // that mirror the real-converter tests still hold against this fake.
        File.WriteAllBytes(pdfPath, System.Text.Encoding.ASCII.GetBytes("%PDF-1.4\n%fake-pdf-for-tests\n"));
        return PdfConversionResult.Success(Name, pdfPath, TimeSpan.Zero);
    }
}
