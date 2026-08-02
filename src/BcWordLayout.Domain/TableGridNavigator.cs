using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Domain;

/// <summary>
/// THE single source of truth for enumerating a table's rows, cells, and grid — <c>gridSpan</c>-aware.
/// <see cref="TableStructureReader"/> (which feeds <c>get_layout_info</c>'s <c>tables[]</c>) and
/// <see cref="LocationResolver"/> (which feeds <c>tableCell</c> addressing) both route their row/cell
/// walking and table-by-index resolution through this class rather than keeping their own copies, so an
/// index a caller reads back from <c>get_layout_info</c> addresses the same physical spot for a structural
/// edit here too, BY CONSTRUCTION rather than by hand-kept parity. This matters because real BC tables are
/// pervasively <c>gridSpan</c>-spanned and ragged (measured across the whole corpus; see
/// <see cref="TableStructureEditor"/>'s remarks for the consequences): a
/// physical cell index is a different visual position in every row, so column-level edits must reason in
/// terms of GRID columns, which is what this navigator exposes.
/// </summary>
/// <remarks>
/// A "row" is a bare <c>w:tr</c> OR a row-level control (<see cref="SdtRow"/>, possibly a
/// <c>repeatingSection</c> wrapping a <c>repeatingSectionItem</c>) descended to the inner <c>w:tr</c> it
/// ultimately holds; a "cell" is a bare <c>w:tc</c> OR a cell-level control (<see cref="SdtCell"/>)
/// descended to the inner <c>w:tc</c> it wraps. <see cref="RowSlot.Wrappers"/> exposes every level of a
/// nested row-control chain (not just the outermost) for the one caller — <see cref="TableStructureReader"/>
/// — that needs to attribute EACH wrapper level to the row's coordinate, not just the outermost id.
/// </remarks>
internal static class TableGridNavigator
{
    /// <summary>
    /// One row of a table in document order. <see cref="RowChild"/> is the direct child of the
    /// <c>w:tbl</c> (a <c>w:tr</c> or an <see cref="SdtRow"/> wrapper); <see cref="InnerRow"/> is the
    /// innermost <c>w:tr</c> the chain of wrappers ultimately holds (null when a row-level control wraps
    /// none). <see cref="IsControlRow"/>/<see cref="ControlId"/> describe the outermost wrapper.
    /// <see cref="Wrappers"/> is the full chain of row-level control wrappers from outermost to innermost
    /// (empty for a bare <c>w:tr</c>) — e.g. <c>[repeatingSection, repeatingSectionItem]</c> — for callers
    /// that must attribute every wrapper level, not just <see cref="RowChild"/>, to this row.
    /// </summary>
    public readonly record struct RowSlot(
        OpenXmlElement RowChild, TableRow? InnerRow, bool IsControlRow, int? ControlId, IReadOnlyList<SdtRow> Wrappers);

    /// <summary>
    /// One cell of a row in document order. <see cref="CellChild"/> is the direct child of the <c>w:tr</c>
    /// (a <c>w:tc</c> or an <see cref="SdtCell"/> wrapper); <see cref="InnerCell"/> is the <c>w:tc</c> it
    /// resolves to (null when a cell-level control wraps none). <see cref="GridSpan"/> is how many grid
    /// columns the cell occupies (from <c>w:tcPr/w:gridSpan</c>, default 1).
    /// </summary>
    public readonly record struct CellSlot(OpenXmlElement CellChild, TableCell? InnerCell, int GridSpan);

    /// <summary>
    /// Every <c>w:tbl</c> in <paramref name="partRoot"/>, in <c>Descendants&lt;Table&gt;()</c> document
    /// order (nested tables flattened into that same list, each getting its own index) — the canonical
    /// table numbering every consumer (<see cref="TableAt"/>, <see cref="TableStructureReader"/>) uses.
    /// </summary>
    public static IReadOnlyList<Table> Tables(OpenXmlElement partRoot)
    {
        ArgumentNullException.ThrowIfNull(partRoot);
        return partRoot.Descendants<Table>().ToList();
    }

    /// <summary>
    /// The <paramref name="tableIndex"/>-th <c>w:tbl</c> in <paramref name="partRoot"/> (see
    /// <see cref="Tables"/> for the ordering rule). Throws a <see cref="NotFoundException"/>
    /// (<see cref="NotFoundTarget.TableCoordinate"/>, surfaced as <c>not_found</c>) when out of range, with
    /// the same wording <see cref="LocationResolver"/> uses.
    /// </summary>
    public static Table TableAt(OpenXmlElement partRoot, int tableIndex, string partDescription)
    {
        var tables = Tables(partRoot);
        if (tableIndex < 0 || tableIndex >= tables.Count)
        {
            throw new NotFoundException(
                $"Table index {tableIndex} is out of range; {partDescription} has {tables.Count} table(s).",
                NotFoundTarget.TableCoordinate);
        }

        return tables[tableIndex];
    }

    /// <summary>The number of grid columns declared by <paramref name="table"/>'s <c>w:tblGrid</c> (0 if absent).</summary>
    public static int GridColumnCount(Table table)
    {
        ArgumentNullException.ThrowIfNull(table);
        return table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().Count() ?? 0;
    }

    /// <summary>The rows of <paramref name="table"/> in document order (see <see cref="RowSlot"/>).</summary>
    public static IReadOnlyList<RowSlot> Rows(Table table)
    {
        ArgumentNullException.ThrowIfNull(table);
        var slots = new List<RowSlot>();
        foreach (var child in table.ChildElements.Where(e => e is TableRow or SdtRow))
        {
            if (child is SdtRow sdtRow)
            {
                var chain = ResolveRowWrapperChain(sdtRow);
                slots.Add(new RowSlot(sdtRow, chain.InnerRow, IsControlRow: true, ControlId: SdtInspector.ReadControlId(sdtRow), Wrappers: chain.Wrappers));
            }
            else
            {
                slots.Add(new RowSlot(child, (TableRow)child, IsControlRow: false, ControlId: null, Wrappers: Array.Empty<SdtRow>()));
            }
        }

        return slots;
    }

    /// <summary>The cells of <paramref name="innerRow"/> in document order (see <see cref="CellSlot"/>).</summary>
    public static IReadOnlyList<CellSlot> Cells(TableRow innerRow)
    {
        ArgumentNullException.ThrowIfNull(innerRow);
        var slots = new List<CellSlot>();
        foreach (var child in innerRow.ChildElements.Where(e => e is TableCell or SdtCell))
        {
            var inner = child switch
            {
                TableCell tc => tc,
                SdtCell sc => sc.SdtContentCell?.GetFirstChild<TableCell>(),
                _ => null,
            };

            slots.Add(new CellSlot(child, inner, GridSpanOf(inner)));
        }

        return slots;
    }

    /// <summary>The <c>w:gridSpan</c> of <paramref name="cell"/> (default 1 for null/absent/&lt;1).</summary>
    public static int GridSpanOf(TableCell? cell)
    {
        var span = cell?.GetFirstChild<TableCellProperties>()?.GetFirstChild<GridSpan>()?.Val?.Value;
        return span is > 1 ? span.Value : 1;
    }

    /// <summary>The leading/trailing skipped-grid-cell counts declared by a row's <c>w:trPr</c> (default 0).</summary>
    public static (int Before, int After) GridEdges(TableRow innerRow)
    {
        ArgumentNullException.ThrowIfNull(innerRow);
        var pr = innerRow.GetFirstChild<TableRowProperties>();
        var before = pr?.GetFirstChild<GridBefore>()?.Val?.Value ?? 0;
        var after = pr?.GetFirstChild<GridAfter>()?.Val?.Value ?? 0;
        return (before, after);
    }

    /// <summary>
    /// How one row maps onto the table's grid: <paramref name="innerRow"/>'s leading/trailing skipped grid
    /// columns and the span its physical cells cover between them.
    /// </summary>
    /// <remarks>
    /// The FIRST grid column a physical cell occupies is <see cref="Before"/>, not 0 — the single fact every
    /// column operation has to respect on a <c>w:gridBefore</c>/<c>w:gridAfter</c> row, and the reason those
    /// rows were rejected outright before this type carried the offset. <see cref="Total"/> is what a
    /// well-formed row's coverage must equal: the table's grid column count (the same identity
    /// <see cref="TableGridConsistencyGuard"/> verifies after every edit).
    /// </remarks>
    public readonly record struct RowGridCoverage(int Before, int SpanTotal, int After)
    {
        /// <summary>Total grid columns this row accounts for — must equal the table's grid column count.</summary>
        public int Total => Before + SpanTotal + After;

        /// <summary>The grid index one past the row's last physical cell (where a trailing skipped region starts).</summary>
        public int ContentEnd => Before + SpanTotal;
    }

    /// <summary>Reads <paramref name="innerRow"/>'s <see cref="RowGridCoverage"/>.</summary>
    public static RowGridCoverage Coverage(TableRow innerRow)
    {
        ArgumentNullException.ThrowIfNull(innerRow);
        var (before, after) = GridEdges(innerRow);
        return new RowGridCoverage(before, Cells(innerRow).Sum(c => c.GridSpan), after);
    }

    /// <summary>
    /// Sets <paramref name="innerRow"/>'s <c>w:gridBefore</c>/<c>w:gridAfter</c> to <paramref name="before"/>/
    /// <paramref name="after"/>, adding the <c>w:trPr</c> or the elements themselves when needed and REMOVING
    /// an element whose new value is 0 (a <c>w:gridAfter w:val="0"</c> is legal but is not how Word writes
    /// "no skipped columns", and leaving one behind would make round-tripped rows differ from authored ones).
    /// </summary>
    /// <remarks>
    /// <c>w:gridBefore</c>/<c>w:gridAfter</c> are the FIRST two elements of the CT_TrPr sequence in that
    /// order, so each is inserted at the front — <c>gridAfter</c> after any existing <c>gridBefore</c>. Order
    /// matters: <c>OpenXmlValidator</c> rejects a mis-sequenced <c>w:trPr</c>, and the pre-save validation
    /// gate would then refuse the whole edit.
    /// </remarks>
    public static void SetGridEdges(TableRow innerRow, int before, int after)
    {
        ArgumentNullException.ThrowIfNull(innerRow);
        ArgumentOutOfRangeException.ThrowIfNegative(before);
        ArgumentOutOfRangeException.ThrowIfNegative(after);

        var pr = innerRow.GetFirstChild<TableRowProperties>();
        if (pr is null)
        {
            if (before == 0 && after == 0)
            {
                return; // Nothing to record and nothing to clear.
            }

            pr = new TableRowProperties();
            innerRow.InsertAt(pr, 0);
        }

        Apply<GridBefore>(pr, before, 0);
        Apply<GridAfter>(pr, after, pr.GetFirstChild<GridBefore>() is null ? 0 : 1);

        static void Apply<T>(TableRowProperties pr, int value, int index)
            where T : OpenXmlLeafElement, new()
        {
            var existing = pr.GetFirstChild<T>();
            if (value == 0)
            {
                existing?.Remove();
                return;
            }

            var element = existing;
            if (element is null)
            {
                element = new T();
                pr.InsertAt(element, Math.Min(index, pr.ChildElements.Count));
            }

            // GridBefore/GridAfter both expose Val as an Int32Value; set it through the concrete property.
            switch (element)
            {
                case GridBefore gb: gb.Val = value; break;
                case GridAfter ga: ga.Val = value; break;
            }
        }
    }

    /// <summary>The full chain of nested row-level control wrappers <see cref="ResolveRowWrapperChain"/> returns.</summary>
    public readonly record struct RowWrapperChain(IReadOnlyList<SdtRow> Wrappers, TableRow? InnerRow);

    /// <summary>
    /// Descends a chain of row-level controls starting at <paramref name="outer"/> (a <c>repeatingSection</c>
    /// may wrap a <c>repeatingSectionItem</c> which wraps the data <c>w:tr</c>), returning every wrapper
    /// visited (outermost first) plus the innermost <c>w:tr</c> the chain ultimately holds (null if it holds
    /// none). Used both by <see cref="Rows"/> (which only needs <see cref="RowWrapperChain.InnerRow"/>) and
    /// by callers that must attribute EVERY wrapper level to the row's coordinate, not just the outermost
    /// one <see cref="RowSlot.RowChild"/> reports (e.g. <see cref="TableStructureReader"/>, whose per-control
    /// coordinate map needs a nested <c>repeatingSectionItem</c>'s own id mapped to its row too).
    /// </summary>
    public static RowWrapperChain ResolveRowWrapperChain(SdtRow outer)
    {
        ArgumentNullException.ThrowIfNull(outer);
        var wrappers = new List<SdtRow>();
        SdtRow? current = outer;
        while (current is not null)
        {
            wrappers.Add(current);
            var content = current.SdtContentRow;
            var innerRow = content?.GetFirstChild<TableRow>();
            if (innerRow is not null)
            {
                return new RowWrapperChain(wrappers, innerRow);
            }

            current = content?.GetFirstChild<SdtRow>();
        }

        return new RowWrapperChain(wrappers, null);
    }
}
