using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Domain;

/// <summary>
/// Works out the run formatting that content inserted INLINE at an <see cref="InsertionAnchor"/> should carry,
/// so a separator, field, label or page number matches the text it lands beside (GitHub issue #122). The rule
/// is Word's own for typing: new text takes the formatting of the character before it, and in an empty
/// paragraph the formatting of the paragraph mark.
/// </summary>
/// <remarks>
/// <para>
/// WHY: every insert used to build a bare <c>w:r</c>, which renders with the paragraph's and the style's
/// defaults. Stock layouts format their runs directly, so the inserted run stood out: StandardSalesQuote's
/// header2 sets <c>w:rStyle w:val="Strong"</c> on every run around <c>DocumentTitle_Lbl</c>/<c>DocumentNo</c>,
/// and an <c>insert_text " / "</c> there came out regular in a bold line. Other stock layouts set Segoe UI
/// 7–8 pt on each run, which no knob could reproduce. Copying the neighbour's <c>w:rPr</c> emits only
/// formatting the file already carries (ADR-0005), the same move <see cref="TableStructureEditor"/>'s
/// <c>insert_column</c> already makes when it builds a cell like its neighbour.
/// </para>
/// <para>
/// WHAT IS NOT COPIED, because it describes the source run's state rather than its look:
/// <list type="bullet">
/// <item>the <c>PlaceholderText</c> character style, which Word puts on a control's content while it shows
/// placeholder text — inheriting it would grey out static text;</item>
/// <item>tracked-change records (<c>w:rPrChange</c>, and the <c>w:ins</c>/<c>w:del</c>/<c>w:moveFrom</c>/
/// <c>w:moveTo</c> marks a paragraph mark can carry);</item>
/// <item>hidden text (<c>w:vanish</c>, <c>w:specVanish</c>, <c>w:webHidden</c>): an insert nobody can see is
/// never what the caller asked for.</item>
/// </list>
/// Picture runs are never a source: a drawing run's <c>w:rPr</c> is not text formatting, and the
/// <c>w:w w:val="0"</c> Word writes there is not even valid on a text run (GitHub issue #54). The search
/// steps back past them to the text before.
/// </para>
/// </remarks>
internal static class InheritedRunFormat
{
    private const string PlaceholderTextStyle = "PlaceholderText";

    /// <summary>
    /// The formatting for content inserted into <paramref name="paragraph"/> right after
    /// <paramref name="insertAfter"/>, or at the paragraph's end when that is null. Returns a fresh, detached
    /// <see cref="RunProperties"/> the caller may clone onto each new run, or null when the text before the
    /// insertion point is unformatted (or there is nothing before it and the paragraph mark is unformatted too).
    /// </summary>
    internal static RunProperties? Before(Paragraph paragraph, OpenXmlElement? insertAfter)
    {
        ArgumentNullException.ThrowIfNull(paragraph);

        if (insertAfter is not null)
        {
            // The anchor is the inline control the caller named (afterControl). A picture control has no text
            // formatting to give, and nothing else is inherited in its place: the caller chose that anchor.
            return TryInline(insertAfter, out var fromAnchor) ? Usable(fromAnchor) : null;
        }

        foreach (var child in paragraph.ChildElements.Reverse())
        {
            if (TryInline(child, out var source))
            {
                return Usable(source);
            }
        }

        return Usable(FromParagraphMark(paragraph));
    }

    /// <summary>
    /// True when <paramref name="element"/> is inline TEXT content (a run, an inline content control, or a
    /// container of runs such as a hyperlink or simple field), with <paramref name="rPr"/> its formatting —
    /// null when that text is unformatted. False for anything that is not text: bookmarks, proofing marks,
    /// the paragraph's own <c>w:pPr</c>, and pictures.
    /// </summary>
    private static bool TryInline(OpenXmlElement element, out RunProperties? rPr)
    {
        rPr = null;
        switch (element)
        {
            case ParagraphProperties:
                return false;

            case Run run:
                if (IsGraphic(run))
                {
                    return false;
                }

                rPr = run.RunProperties;
                return true;

            case SdtRun sdt:
                if (SdtInspector.Classify(sdt.SdtProperties) == SdtInspector.Classification.Picture
                    || sdt.Descendants<Drawing>().Any())
                {
                    return false;
                }

                // The control's own w:sdtPr/w:rPr is the formatting of the VALUE BC renders into it; its content
                // run may still be showing placeholder text. Stock controls carry both, identically.
                rPr = sdt.SdtProperties?.GetFirstChild<RunProperties>() is { HasChildren: true } own
                    ? own
                    : LastTextRun(sdt)?.RunProperties;
                return true;

            default:
                if (LastTextRun(element) is { } last)
                {
                    rPr = last.RunProperties;
                    return true;
                }

                return false;
        }
    }

    private static Run? LastTextRun(OpenXmlElement container) =>
        container.Descendants<Run>().LastOrDefault(r => !IsGraphic(r));

    /// <summary>True for a run that carries a picture or embedded object rather than text.</summary>
    internal static bool IsGraphic(Run run) =>
        run.GetFirstChild<Drawing>() is not null
        || run.GetFirstChild<Picture>() is not null
        || run.GetFirstChild<EmbeddedObject>() is not null;

    /// <summary>
    /// The paragraph mark's run properties (<c>w:pPr/w:rPr</c>) re-typed as <see cref="RunProperties"/>. The
    /// two share their formatting children, but the mark's own revision marks (<c>w:ins</c>, <c>w:del</c>,
    /// <c>w:moveFrom</c>, <c>w:moveTo</c>, <c>w:rPrChange</c>) are not valid in a run's <c>w:rPr</c>, so
    /// <see cref="OpenXmlCompositeElement.AddChild"/> — which places each child in schema order and refuses
    /// one the target cannot hold — drops exactly those.
    /// </summary>
    private static RunProperties? FromParagraphMark(Paragraph paragraph)
    {
        if (paragraph.ParagraphProperties?.ParagraphMarkRunProperties is not { HasChildren: true } mark)
        {
            return null;
        }

        var rPr = new RunProperties();
        foreach (var child in mark.ChildElements)
        {
            rPr.AddChild(child.CloneNode(true), throwOnError: false);
        }

        return rPr;
    }

    /// <summary>
    /// A detached copy of <paramref name="source"/> without the state-describing children this class does not
    /// copy (see the remarks), or null when nothing is left.
    /// </summary>
    private static RunProperties? Usable(RunProperties? source)
    {
        if (source is null)
        {
            return null;
        }

        var rPr = (RunProperties)source.CloneNode(true);
        if (rPr.RunStyle?.Val?.Value is PlaceholderTextStyle)
        {
            rPr.RunStyle = null;
        }

        rPr.RemoveAllChildren<RunPropertiesChange>();
        rPr.RemoveAllChildren<Vanish>();
        rPr.RemoveAllChildren<SpecVanish>();
        rPr.RemoveAllChildren<WebHidden>();

        return rPr.HasChildren ? rPr : null;
    }
}
