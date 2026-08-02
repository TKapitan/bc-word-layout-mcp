namespace BcWordLayout.Render;

/// <summary>Discovers and selects the available <see cref="IPdfConverter"/> implementations.</summary>
public static class PdfConverterFactory
{
    /// <summary>
    /// Every converter this build knows how to construct, most-preferred first: Word (only on Windows,
    /// since <see cref="WordComConverter"/> is a Windows-only type), then LibreOffice. Each instance's
    /// <see cref="IPdfConverter.IsAvailable"/> reflects live machine state at the time it is read.
    /// </summary>
    public static IReadOnlyList<IPdfConverter> All()
    {
        var converters = new List<IPdfConverter>();
        if (OperatingSystem.IsWindows())
        {
            converters.Add(new WordComConverter());
        }

        converters.Add(new LibreOfficeConverter());
        return converters;
    }

    /// <summary>Whether at least one converter is currently available on this machine.</summary>
    public static bool AnyAvailable => All().Any(c => c.IsAvailable);

    /// <summary>
    /// Resolves <paramref name="preference"/> to a concrete converter.
    /// <see cref="PdfConverterKind.Auto"/> prefers Word when available, falls back to LibreOffice when
    /// available, and otherwise returns a LibreOffice instance whose <see cref="IPdfConverter.Convert"/>
    /// will cleanly fail (never throw) with an actionable "no PDF converter available" error. Asking for a
    /// specific converter always returns a converter of that kind, even when its dependency is missing —
    /// its own <c>Convert</c> then fails cleanly instead of the factory throwing or returning null.
    /// </summary>
    public static IPdfConverter Select(PdfConverterKind preference)
    {
        switch (preference)
        {
            case PdfConverterKind.Word:
                // Guarded: WordComConverter is [SupportedOSPlatform("windows")], so it may only be
                // constructed behind an OperatingSystem.IsWindows() check.
                if (OperatingSystem.IsWindows())
                {
                    return new WordComConverter();
                }

                return new NotSupportedPdfConverter
                {
                    Name = "word-com",
                    Reason = "Word COM automation is only available on Windows.",
                };

            case PdfConverterKind.LibreOffice:
                return new LibreOfficeConverter();

            case PdfConverterKind.Auto:
            default:
                if (OperatingSystem.IsWindows())
                {
                    var word = new WordComConverter();
                    if (word.IsAvailable)
                    {
                        return word;
                    }
                }

                // LibreOfficeConverter fails cleanly on its own when soffice isn't installed either, which
                // is exactly the "no PDF converter available" behavior Auto needs as its last resort.
                return new LibreOfficeConverter();
        }
    }
}

/// <summary>
/// Stand-in returned by <see cref="PdfConverterFactory.Select"/> for a converter kind that cannot even be
/// constructed on this OS (currently: Word requested on non-Windows). <see cref="IsAvailable"/> is always
/// false and <see cref="Convert"/> always fails cleanly with <see cref="Reason"/>, so callers never need to
/// special-case "this converter doesn't exist on this platform" versus "it exists but its dependency is
/// missing" — both simply come back as an unavailable converter whose <c>Convert</c> fails cleanly.
/// </summary>
internal sealed class NotSupportedPdfConverter : IPdfConverter
{
    public required string Name { get; init; }

    public required string Reason { get; init; }

    public bool IsAvailable => false;

    public PdfConversionResult Convert(string docxPath, string pdfPath, PdfConversionOptions? options = null) =>
        PdfConversionResult.Failure(Name, Reason);
}
