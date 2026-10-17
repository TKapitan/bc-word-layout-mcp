using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Domain;

/// <summary>
/// The package-wide view of content-control ids (<c>w:sdtPr/w:id</c>) across the document body, headers and
/// footers — what validate_layout's <c>duplicate-control-id</c> warning reports, and what every id-addressed
/// tool (remove_control, an <c>afterControl</c> location, insert_repeater_row/insert_subtotal_row's
/// parentRepeaterId) checks before acting on an id.
/// </summary>
/// <remarks>
/// Those lookups used to take the FIRST control carrying the id, in part order (GitHub issue #101): on a copy
/// of JobQuote.docx whose header2 control had been given the id of a body control, remove_control removed the
/// body control without a word, and validation reported nothing at all. An id carried twice does not say
/// which control is meant, so the lookups now refuse it (ADR-0003). Uniqueness is package-wide, not per part,
/// even for an <c>afterControl</c> location that only searches one part: an agent that read the id off a
/// header control and forgot <c>layoutPart</c> would otherwise land silently on the body control instead of
/// getting a not_found that tells it to look again.
/// </remarks>
internal static class ControlIds
{
    /// <summary>Every control carrying <paramref name="controlId"/>, with its part, in body/header/footer order.</summary>
    internal static IReadOnlyList<(SdtElement Control, string Part)> FindAll(MainDocumentPart main, int controlId) =>
        PartWalker.ContentParts(main)
            .SelectMany(p => p.Root.Descendants<SdtElement>().Select(sdt => (Control: sdt, p.PartName)))
            .Where(c => SdtInspector.ReadControlId(c.Control) == controlId)
            .Select(c => (c.Control, c.PartName))
            .ToList();

    /// <summary>
    /// Throws the <see cref="LayoutDefect.DuplicateControlId"/> refusal when more than one control carries
    /// <paramref name="controlId"/>; a no-op otherwise (an id carried by no control is the caller's own
    /// not_found to report).
    /// </summary>
    internal static void EnsureUnique(MainDocumentPart main, int controlId)
    {
        var carriers = FindAll(main, controlId);
        if (carriers.Count > 1)
        {
            throw DuplicateDefect(controlId, carriers);
        }
    }

    /// <summary>The refusal for an id more than one control carries, naming every one of them.</summary>
    internal static LayoutDefectException DuplicateDefect(int controlId, IReadOnlyList<(SdtElement Control, string Part)> carriers) =>
        new($"Control id {controlId.ToString(CultureInfo.InvariantCulture)} is carried by {carriers.Count} content "
            + $"controls in this layout ({Describe(carriers)}), so it does not identify one control; the edit was "
            + "not made rather than applied to whichever comes first.",
            LayoutDefect.DuplicateControlId);

    /// <summary>Every id carried by more than one control, with its carriers, in first-occurrence order.</summary>
    internal static IEnumerable<(int Id, IReadOnlyList<(SdtElement Control, string Part)> Carriers)> Duplicates(MainDocumentPart main) =>
        PartWalker.ContentParts(main)
            .SelectMany(p => p.Root.Descendants<SdtElement>().Select(sdt => (Control: sdt, p.PartName, Id: SdtInspector.ReadControlId(sdt))))
            .Where(c => c.Id.HasValue)
            .GroupBy(c => c.Id!.Value)
            .Where(g => g.Count() > 1)
            .Select(g => (g.Key, (IReadOnlyList<(SdtElement Control, string Part)>)g.Select(c => (c.Control, c.PartName)).ToList()));

    /// <summary>
    /// The carriers as "part: alias" (or the tag, the binding XPath, or the control kind for a control with
    /// none of those), e.g. <c>document.xml: '#Nav: /Job/BillToAddress3', header2.xml: unbound control</c>.
    /// </summary>
    internal static string Describe(IReadOnlyList<(SdtElement Control, string Part)> carriers) =>
        string.Join(
            ", ",
            carriers.Select(c =>
            {
                var name = SdtInspector.ReadAlias(c.Control) ?? SdtInspector.ReadTag(c.Control) ?? SdtInspector.ReadXPath(c.Control);
                return name is null
                    ? $"{c.Part}: {SdtInspector.ClassifyControlKind(c.Control).ToString().ToLowerInvariant()} control"
                    : $"{c.Part}: '{name}'";
            }));
}
