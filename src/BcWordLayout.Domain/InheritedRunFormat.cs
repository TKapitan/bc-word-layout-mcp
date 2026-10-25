using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Domain;

/// <summary>
/// Works out the run formatting that content inserted INLINE at an <see cref="InsertionAnchor"/> should carry,
/// so a separator, field, label or page number matches the text it lands beside (GitHub issue #122). The rule
/// is Word's own for typing: new text takes the formatting of the character before it, and in an empty
/// paragraph the formatting of the paragraph mark. The cells of a new table row take theirs from the row beside
/// it, column by column (<see cref="ForRowCells"/>), and the paragraph style with it
/// (<see cref="ParagraphStylesForRowCells"/>).
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
    /// The formatting each cell of a NEW table row takes from <paramref name="reference"/>, the row it is
    /// inserted next to — Word's rule for an inserted row, which keeps the formatting of the row it was inserted
    /// from. For each entry of <paramref name="spans"/> (the new row's cells, left to right from grid column 0),
    /// the formatting of the reference cell that covers the new cell's first grid column (the nearest cell when
    /// the reference row skips that column), read in that cell's first paragraph holding text exactly as
    /// <see cref="Before"/> reads a paragraph's end, else from its first paragraph's mark. An entry is null where
    /// that cell's text is unformatted, and every entry is null when <paramref name="reference"/> is.
    /// </summary>
    /// <remarks>
    /// WHY per column: a stock row formats each cell on its own — <c>SalespersonCommission.docx</c> (115) and
    /// <c>JobQuote.docx</c> (1016) set Segoe UI 7–8 pt on every run and paragraph mark of their rows, and
    /// InventoryOrderDetails mixes bold and regular cells in one row. A row the tools built bare printed at the
    /// style default (11 pt) inside such a table (sandbox pack P04, BC 28). A table whose rows carry no direct
    /// formatting (the stock quote's lines table, every tool-built table) gives nothing, and the new row looks
    /// like its neighbours through the same styles they use.
    /// </remarks>
    internal static IReadOnlyList<RunProperties?> ForRowCells(TableRow? reference, IReadOnlyList<int> spans) =>
        ReferenceCells(reference, spans).Select(cell => cell is null ? null : OfCell(cell)).ToList();

    /// <summary>
    /// The paragraph style each cell of a new table row takes from <paramref name="reference"/>: per entry of
    /// <paramref name="spans"/>, the <c>w:pStyle</c> of the paragraph <see cref="ForRowCells"/> reads in the
    /// same reference cell. An entry is null where that paragraph names no style, and every entry is null when
    /// <paramref name="reference"/> is.
    /// </summary>
    /// <remarks>
    /// WHY: some stock tables size their text through paragraph styles rather than on the runs.
    /// StandardSalesInvoiceVatSpec's line row sets <c>LeftAlign</c>/<c>RightAlign</c> on each cell
    /// (<c>RightAlign</c> carries the 10 pt), so a row built beside it printed a point larger with the run
    /// formatting alone (sandbox pack P22, BC 28).
    /// </remarks>
    internal static IReadOnlyList<string?> ParagraphStylesForRowCells(TableRow? reference, IReadOnlyList<int> spans) =>
        ReferenceCells(reference, spans)
            .Select(cell => cell is null ? null : TextParagraphOf(cell)?.ParagraphProperties?.ParagraphStyleId?.Val?.Value)
            .ToList();

    /// <summary>
    /// The cell of <paramref name="reference"/> each new cell takes its formatting from (see
    /// <see cref="ForRowCells"/>), or null for every new cell when there is no reference row with cells.
    /// </summary>
    private static OpenXmlElement?[] ReferenceCells(TableRow? reference, IReadOnlyList<int> spans)
    {
        ArgumentNullException.ThrowIfNull(spans);

        var sources = new OpenXmlElement?[spans.Count];
        var cells = reference is null ? [] : TableGridNavigator.Cells(reference);
        if (cells.Count == 0)
        {
            return sources;
        }

        var firstColumn = TableGridNavigator.Coverage(reference!).Before;
        var gridColumn = 0;
        for (var i = 0; i < spans.Count; i++)
        {
            var source = CellCovering(cells, firstColumn, gridColumn);
            sources[i] = source.InnerCell ?? source.CellChild;
            gridColumn += spans[i];
        }

        return sources;
    }

    /// <summary>
    /// The cell of <paramref name="cells"/> (whose first one starts at grid column <paramref name="firstColumn"/>)
    /// covering <paramref name="gridColumn"/>; the first cell for a column before it, the last for one after.
    /// </summary>
    private static TableGridNavigator.CellSlot CellCovering(
        IReadOnlyList<TableGridNavigator.CellSlot> cells, int firstColumn, int gridColumn)
    {
        var start = firstColumn;
        foreach (var cell in cells)
        {
            if (gridColumn < start + cell.GridSpan)
            {
                return cell;
            }

            start += cell.GridSpan;
        }

        return cells[^1];
    }

    /// <summary>
    /// The formatting of <paramref name="cell"/>'s text, as <see cref="ForRowCells"/> reads it: at the end of its
    /// text paragraph (<see cref="TextParagraphOf"/>), whose own mark gives it when the paragraph is empty.
    /// </summary>
    private static RunProperties? OfCell(OpenXmlElement cell) =>
        TextParagraphOf(cell) is { } paragraph ? Before(paragraph, insertAfter: null) : null;

    /// <summary>The first paragraph of <paramref name="cell"/> holding text, else its first paragraph, else null.</summary>
    private static Paragraph? TextParagraphOf(OpenXmlElement cell)
    {
        Paragraph? first = null;
        foreach (var paragraph in cell.Descendants<Paragraph>())
        {
            first ??= paragraph;
            if (paragraph.ChildElements.Any(child => TryInline(child, out _)))
            {
                return paragraph;
            }
        }

        return first;
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
