using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace BcWordLayout.Domain;

/// <summary>
/// THE single source of truth for "walk the main document body plus every header and footer part, in that
/// order" — a hand-written three-block pattern (main document, then <c>foreach header</c>, then
/// <c>foreach footer</c>, each guarding a possibly-null content root) that had been copy-pasted, together
/// with its own <c>PartFileName</c> helper, into roughly eight call sites across <c>LayoutReader</c>,
/// <c>LayoutEditor</c>, <c>LayoutValidator</c>, <c>LayoutRefresher</c>, <c>LocationResolver</c>,
/// <c>TableGridConsistencyGuard</c>, <c>PlainTextNestingGuard</c>, and <c>BcWordLayout.Merge.MergeEngine</c>
/// Every caller wants the SAME three-part order
/// (document.xml, then headers, then footers) and the SAME "skip a part with no content root" behavior;
/// this type is the one place that ordering and null-skipping is expressed.
/// </summary>
internal static class PartWalker
{
    /// <summary>The OOXML part file name (e.g. <c>document.xml</c>, <c>header1.xml</c>) of <paramref name="part"/>.</summary>
    internal static string PartFileName(OpenXmlPart part) => Path.GetFileName(part.Uri.OriginalString);

    /// <summary>
    /// Yields (content root, part file name) for the main document plus every header and footer part of
    /// <paramref name="main"/>, in that order — document.xml first, then headers, then footers, each in
    /// their own collection order — skipping any part whose content root is null (no content to walk).
    /// </summary>
    internal static IEnumerable<(OpenXmlPartRootElement Root, string PartName)> ContentParts(MainDocumentPart main)
    {
        if (main.Document is not null)
        {
            yield return (main.Document, "document.xml");
        }

        foreach (var header in main.HeaderParts)
        {
            if (header.Header is not null)
            {
                yield return (header.Header, PartFileName(header));
            }
        }

        foreach (var footer in main.FooterParts)
        {
            if (footer.Footer is not null)
            {
                yield return (footer.Footer, PartFileName(footer));
            }
        }
    }

    /// <summary>
    /// Same order and null-skipping as <see cref="ContentParts"/>, but also yields the owning
    /// <see cref="OpenXmlPart"/> itself — for callers (e.g. <c>MergeEngine</c>'s picture fill, which must
    /// add a placeholder <c>ImagePart</c> to whichever main/header/footer part hosts the picture control it
    /// found) that need to act on the part, not just read its content root.
    /// </summary>
    internal static IEnumerable<(OpenXmlPartRootElement Root, OpenXmlPart Part, string PartName)> ContentPartsWithHost(MainDocumentPart main)
    {
        if (main.Document is not null)
        {
            yield return (main.Document, main, "document.xml");
        }

        foreach (var header in main.HeaderParts)
        {
            if (header.Header is not null)
            {
                yield return (header.Header, header, PartFileName(header));
            }
        }

        foreach (var footer in main.FooterParts)
        {
            if (footer.Footer is not null)
            {
                yield return (footer.Footer, footer, PartFileName(footer));
            }
        }
    }
}
