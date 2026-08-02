using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Domain;

/// <summary>
/// One content control found nested inside a PLAIN-TEXT content control — a combination Word rejects as
/// a corrupt document even though it is well-formed OOXML that <see cref="DocumentFormat.OpenXml.Validation.OpenXmlValidator"/>
/// accepts (see <see cref="PlainTextNestingGuard"/>). Carries enough identity (alias/id of both the inner
/// control and its offending plain-text ancestor, plus the part) to build an agent-actionable rejection
/// message.
/// </summary>
public readonly record struct PlainTextNesting(
    string? InnerAlias, int? InnerId, string? OuterAlias, int? OuterId, string Part)
{
    /// <summary>
    /// A stable, human-readable one-line description. Stable is important: <c>GuardMutate</c> diffs the
    /// BEFORE-edit set against the AFTER-edit set BY THIS STRING, so only nestings this edit newly
    /// introduced are rejected (a layout that somehow already had one stays editable, exactly as the
    /// OpenXmlValidator baseline-diff does for pre-existing structural errors).
    /// </summary>
    public string Describe() =>
        $"content control {Identify(InnerAlias, InnerId)} is nested inside plain-text content control "
        + $"{Identify(OuterAlias, OuterId)} in {Part}";

    private static string Identify(string? alias, int? id) =>
        alias is not null
            ? $"'{alias}' (id {id?.ToString(CultureInfo.InvariantCulture) ?? "?"})"
            : $"id {id?.ToString(CultureInfo.InvariantCulture) ?? "?"}";
}

/// <summary>
/// Detects content controls nested inside a PLAIN-TEXT content control — i.e. an <see cref="SdtElement"/>
/// that has an ancestor <see cref="SdtElement"/> whose <c>w:sdtPr</c> carries a <c>w:text</c>
/// (<see cref="SdtContentText"/>) marker.
/// </summary>
/// <remarks>
/// WHY THIS EXISTS: a plain-text content control (the shape BC uses for most header fields, e.g. the
/// cell-level <c>DocumentNo</c>/address controls — <c>w:sdtPr</c> containing <c>&lt;w:text/&gt;</c>) may
/// contain ONLY runs of text. Word forbids any nested content control inside one and reports the whole
/// document as corrupted ("The file appears to be corrupted") when it finds one — refusing to open it
/// without a repair that strips the offending controls. Crucially this is a Word-enforced SEMANTIC rule,
/// NOT an OOXML schema rule: the nesting is perfectly well-formed markup, so
/// <see cref="DocumentFormat.OpenXml.Validation.OpenXmlValidator"/> accepts it and the pre-save structural
/// gate in <c>GuardMutate</c> never saw it. That is exactly how a single "insert this field after that
/// one" edit could both report success AND leave a file Word cannot open: inserting a field/label
/// <see cref="LocationKind.AfterControl"/> a cell-level plain-text control (or via
/// <see cref="LocationKind.TableCell"/>/<see cref="LocationKind.AtText"/> that lands inside one) anchors
/// the new control INSIDE that plain-text control's own content. This guard is the missing check: it lets
/// <c>GuardMutate</c> reject such an edit before it ever reaches disk, the same way the OpenXmlValidator
/// gate rejects genuinely malformed markup.
/// </remarks>
public static class PlainTextNestingGuard
{
    /// <summary>
    /// Returns every content control anywhere in <paramref name="doc"/> (main document body, then each
    /// header, then each footer — the same part order <c>LayoutEditor</c>'s own control scan uses) that
    /// sits inside a plain-text content control. Empty when the document is clean.
    /// </summary>
    public static IReadOnlyList<PlainTextNesting> Find(WordprocessingDocument doc)
    {
        ArgumentNullException.ThrowIfNull(doc);

        var main = doc.MainDocumentPart
            ?? throw new InvalidDataException("Layout has no main document part.");

        var found = new List<PlainTextNesting>();

        foreach (var (root, part) in PartWalker.ContentParts(main))
        {
            foreach (var inner in root.Descendants<SdtElement>())
            {
                var plainTextAncestor = inner.Ancestors<SdtElement>().FirstOrDefault(IsPlainText);
                if (plainTextAncestor is not null)
                {
                    found.Add(new PlainTextNesting(
                        InnerAlias: SdtInspector.ReadAlias(inner),
                        InnerId: SdtInspector.ReadControlId(inner),
                        OuterAlias: SdtInspector.ReadAlias(plainTextAncestor),
                        OuterId: SdtInspector.ReadControlId(plainTextAncestor),
                        Part: part));
                }
            }
        }

        return found;
    }

    /// <summary>
    /// True when <paramref name="sdt"/> is a plain-text content control: its <c>w:sdtPr</c> carries a
    /// <c>w:text</c> marker (<see cref="SdtContentText"/>), whether single-line (<c>&lt;w:text/&gt;</c>) or
    /// multi-line (<c>&lt;w:text w:multiLine="1"/&gt;</c>) — both equally forbid nested content controls.
    /// </summary>
    private static bool IsPlainText(SdtElement sdt) =>
        sdt.GetFirstChild<SdtProperties>()?.GetFirstChild<SdtContentText>() is not null;
}
