using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Domain;

/// <summary>
/// One disallowed element found inside a PLAIN-TEXT content control — a nested content control, or non-run
/// content (a field code, a table, a drawing) — a combination Word rejects as a corrupt document even though
/// it is well-formed OOXML that <see cref="DocumentFormat.OpenXml.Validation.OpenXmlValidator"/> accepts (see
/// <see cref="PlainTextNestingGuard"/>). Carries enough identity (alias/id of the inner control — or a short
/// name of the inner element — and of its offending plain-text ancestor, plus the part) to build an
/// agent-actionable rejection message.
/// </summary>
/// <param name="InnerAlias">The nested control's alias; null for non-control content.</param>
/// <param name="InnerId">The nested control's <c>w:id</c>; null for non-control content.</param>
/// <param name="OuterAlias">The plain-text ancestor's alias.</param>
/// <param name="OuterId">The plain-text ancestor's <c>w:id</c>.</param>
/// <param name="Part">The part file name both live in.</param>
/// <param name="InnerElement">
/// Non-control content only: a short description of the offending element (e.g. <c>a field code
/// (w:fldChar)</c>). Null when the offender is a nested content control.
/// </param>
public readonly record struct PlainTextNesting(
    string? InnerAlias, int? InnerId, string? OuterAlias, int? OuterId, string Part, string? InnerElement = null)
{
    /// <summary>
    /// A stable, human-readable one-line description. Stable is important: <c>GuardMutate</c> diffs the
    /// BEFORE-edit set against the AFTER-edit set BY THIS STRING (as a multiset), so only nestings this edit
    /// newly introduced are rejected (a layout that somehow already had one stays editable, exactly as the
    /// OpenXmlValidator baseline-diff does for pre-existing structural errors).
    /// </summary>
    public string Describe() =>
        (InnerElement is null
            ? $"content control {Identify(InnerAlias, InnerId)} is nested"
            : $"{InnerElement} sits")
        + $" inside plain-text content control {Identify(OuterAlias, OuterId)} in {Part}";

    private static string Identify(string? alias, int? id) =>
        alias is not null
            ? $"'{alias}' (id {id?.ToString(CultureInfo.InvariantCulture) ?? "?"})"
            : $"id {id?.ToString(CultureInfo.InvariantCulture) ?? "?"}";
}

/// <summary>
/// Detects content Word forbids inside a PLAIN-TEXT content control — an <see cref="SdtElement"/> whose
/// <c>w:sdtPr</c> carries a <c>w:text</c> (<see cref="SdtContentText"/>) marker: a nested content control, or
/// non-run content — a field code (<c>w:fldChar</c>/<c>w:fldSimple</c>), a <c>w:tbl</c>, or a
/// <c>w:drawing</c>.
/// </summary>
/// <remarks>
/// WHY THIS EXISTS: a plain-text content control (the shape BC uses for most header fields, e.g. the
/// cell-level <c>DocumentNo</c>/address controls — <c>w:sdtPr</c> containing <c>&lt;w:text/&gt;</c>) may
/// contain ONLY runs of text. Word forbids a nested content control, a field code, a table or a drawing
/// inside one and reports the whole document as corrupted ("The file appears to be corrupted") when it finds
/// one — refusing to open it without a repair that strips the offending content (reproduced for nested
/// controls, and for <c>insert_page_number</c>'s field runs and <c>insert_table</c>, GitHub issue #76).
/// Crucially this is a Word-enforced SEMANTIC rule, NOT an OOXML schema rule: the markup is perfectly
/// well-formed, so <see cref="DocumentFormat.OpenXml.Validation.OpenXmlValidator"/> accepts it and the
/// pre-save structural gate in <c>GuardMutate</c> never saw it. That is exactly how a single "insert this
/// after that one" edit could both report success AND leave a file Word cannot open.
/// <para>
/// Since issue #76, <see cref="LocationResolver.Resolve"/> refuses any location inside a data-bound
/// control's content up front, so no insert tool reaches this state by that route any more. This guard is the
/// mechanical BACKSTOP behind that refusal: it lets <c>GuardMutate</c> reject such an edit before it ever
/// reaches disk whatever route produced it (including an unbound <c>w:text</c> control, which the resolver
/// deliberately lets through), the same way the OpenXmlValidator gate rejects genuinely malformed markup. No
/// corpus layout carries any of these elements inside a <c>w:text</c> control.
/// </para>
/// </remarks>
public static class PlainTextNestingGuard
{
    /// <summary>
    /// Returns every content control, and every field code / table / drawing, anywhere in
    /// <paramref name="doc"/> (main document body, then each header, then each footer — the same part order
    /// <c>LayoutEditor</c>'s own control scan uses) that sits inside a plain-text content control. Empty when
    /// the document is clean.
    /// </summary>
    public static IReadOnlyList<PlainTextNesting> Find(WordprocessingDocument doc)
    {
        ArgumentNullException.ThrowIfNull(doc);

        var main = doc.MainDocumentPart
            ?? throw new InvalidDataException("Layout has no main document part.");

        var found = new List<PlainTextNesting>();

        foreach (var (root, part) in PartWalker.ContentParts(main))
        {
            foreach (var element in root.Descendants())
            {
                if (!IsForbiddenInPlainText(element, out var innerElement))
                {
                    continue;
                }

                var plainTextAncestor = element.Ancestors<SdtElement>().FirstOrDefault(IsPlainText);
                if (plainTextAncestor is null)
                {
                    continue;
                }

                var inner = element as SdtElement;
                found.Add(new PlainTextNesting(
                    InnerAlias: inner is null ? null : SdtInspector.ReadAlias(inner),
                    InnerId: inner is null ? null : SdtInspector.ReadControlId(inner),
                    OuterAlias: SdtInspector.ReadAlias(plainTextAncestor),
                    OuterId: SdtInspector.ReadControlId(plainTextAncestor),
                    Part: part,
                    InnerElement: innerElement));
            }
        }

        return found;
    }

    /// <summary>
    /// True for an element a plain-text control may not contain; <paramref name="innerElement"/> is its
    /// short description for <see cref="PlainTextNesting.Describe"/>, or null for a nested content control
    /// (described by its own alias/id instead). <c>w:instrText</c> is not listed separately: it only ever
    /// appears between the <c>w:fldChar</c> markers that are.
    /// </summary>
    private static bool IsForbiddenInPlainText(OpenXmlElement element, out string? innerElement)
    {
        innerElement = element switch
        {
            FieldChar => "a field code (w:fldChar)",
            SimpleField => "a field code (w:fldSimple)",
            Table => "a table (w:tbl)",
            Drawing => "a drawing (w:drawing)",
            _ => null,
        };

        return innerElement is not null || element is SdtElement;
    }

    /// <summary>
    /// True when <paramref name="sdt"/> is a plain-text content control: its <c>w:sdtPr</c> carries a
    /// <c>w:text</c> marker (<see cref="SdtContentText"/>), whether single-line (<c>&lt;w:text/&gt;</c>) or
    /// multi-line (<c>&lt;w:text w:multiLine="1"/&gt;</c>) — both equally forbid nested content.
    /// </summary>
    private static bool IsPlainText(SdtElement sdt) =>
        sdt.GetFirstChild<SdtProperties>()?.GetFirstChild<SdtContentText>() is not null;
}
