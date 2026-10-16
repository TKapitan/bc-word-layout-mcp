using System.Globalization;
using BcWordLayout.Domain.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// CA2208 wants ArgumentException.ParamName to be a C# parameter of the throwing method. Here it
// deliberately is not: ParamName carries the MCP TOOL argument name ("edges", "size", "dataPath",
// "atColumn") because ToolGuards.Guard keys the agent-facing hint map on it (InvalidArgumentHint),
// and the tool argument is what the calling agent must fix — the C# parameter is an implementation
// detail it never sees. InvalidArgumentHintCoverageTests pins every one of these hint branches.
#pragma warning disable CA2208

namespace BcWordLayout.Domain;

/// <summary>
/// Deterministic OOXML mutations for a table's STRUCTURE — column widths (<see cref="SetColumnWidths"/>),
/// adding a column (<see cref="InsertColumn"/>), and removing a column (<see cref="RemoveColumn"/>). The
/// third editor in the family alongside <see cref="LayoutEditor"/> (content controls by <c>w:id</c>) and
/// <see cref="CellTextEditor"/> (a single plain-text cell by coordinates); like both, every method works
/// against an ALREADY-OPEN document and does no file I/O — opening, pre-save validation (including the
/// <see cref="TableGridConsistencyGuard"/> and <see cref="OpenXmlValidator"/> gates), and saving are the
/// caller's job (see <c>BcWordLayout.McpHost.Tools.ToolGuards.GuardTableEdit</c>, used by
/// <c>BcWordLayout.McpHost.Tools.TableTools</c>).
/// </summary>
/// <remarks>
/// <para>
/// GRID-COLUMN ADDRESSING. Real BC tables are pervasively <c>gridSpan</c>-spanned and ragged (a physical
/// cell index is a different visual position in every row), so columns are addressed by GRID column index (0..gridCount-1 from <c>w:tblGrid</c>), and every
/// operation reasons per row via <see cref="TableGridNavigator"/>, summing each cell's <c>gridSpan</c> to
/// map physical cells to the grid.
/// </para>
/// <para>
/// SAFETY. Every edit here keeps each row's cell coverage equal to the grid column count; the
/// <see cref="TableGridConsistencyGuard"/> (diffed BEFORE/AFTER in <c>GuardMutate</c>) mechanically rejects
/// any edit that would desync them, so a grid/cell mismatch can never reach disk — the corruption class the
/// <see cref="OpenXmlValidator"/> gate cannot see. Tables using <c>w:vMerge</c> (vertical merges) are
/// rejected up front: that shape is absent from every layout reviewed and its span arithmetic is deferred
/// pending real capture (plan §5-B).
/// </para>
/// <para>
/// SKIPPED GRID COLUMNS. Rows using <c>w:gridBefore</c>/<c>w:gridAfter</c> cover fewer physical cells than
/// the table has grid columns, so their first cell starts at grid column <c>Before</c> rather than 0 and a
/// grid position may fall in a run no cell reaches at all. Every operation here carries that offset via
/// <see cref="TableGridNavigator.RowGridCoverage"/>, and insert/remove grow or shrink the skipped run itself
/// when the target column lands inside one. This shape was rejected outright until 2026-08-01 on the grounds
/// that it was absent from the corpus; it is not — <c>gridAfter</c> appears in seven base-app layouts,
/// including on the line-items repeater table of <c>StandardSalesInvoiceVatSpec.docx</c>, so the rejection was
/// blocking edits to the lines table of a stock BC sales invoice.
/// </para>
/// </remarks>
public static class TableStructureEditor
{
    private const int DefaultColumnWidthTwips = 2000;

    /// <summary>
    /// Rewrites the widths of every grid column of the <paramref name="tableIndex"/>-th table to
    /// <paramref name="widths"/> (one twip value per grid column), then sets each physical cell's
    /// <c>w:tcW</c> to the sum of the grid-column widths that cell spans — so a spanned cell stays as wide
    /// as the columns beneath it. Deterministic on spanned/ragged tables; changes no structure.
    /// </summary>
    /// <exception cref="ArgumentException"><paramref name="widths"/>'s count ≠ the grid column count, or the table uses vMerge.</exception>
    /// <exception cref="NotFoundException"><paramref name="tableIndex"/> is out of range for the part.</exception>
    public static TableEditResult SetColumnWidths(
        WordprocessingDocument doc, LayoutPart part, string? partName, int tableIndex, IReadOnlyList<int> widths)
    {
        ArgumentNullException.ThrowIfNull(doc);
        ArgumentNullException.ThrowIfNull(widths);

        var (table, partFile, gridCount) = ResolveTable(doc, part, partName, tableIndex);
        if (widths.Count != gridCount)
        {
            throw new ArgumentException(
                $"widths has {widths.Count} entr{(widths.Count == 1 ? "y" : "ies")} but table {tableIndex} has "
                + $"{gridCount} grid column(s); supply exactly one width (in twips) per grid column.",
                nameof(widths));
        }

        RejectUnsupportedShape(table, "set_column_widths");

        var rowsAffected = ApplyGridWidths(table, widths);

        SyncTableWidthToGrid(table);

        return new TableEditResult
        {
            Operation = "set_column_widths",
            Part = partFile,
            TableIndex = tableIndex,
            ColumnIndex = null,
            RowsAffected = rowsAffected,
            ColumnCountBefore = gridCount,
            ColumnCountAfter = gridCount,
            Summary = $"Set the {gridCount} column width(s) of table {tableIndex}"
                + (partFile == "document.xml" ? string.Empty : $" in {partFile}")
                + $" and resized {rowsAffected} row(s)' cells to match.",
        };
    }

    /// <summary>
    /// Draws (or clears) per-cell rules on ONE row of the <paramref name="tableIndex"/>-th table: the cell at
    /// <paramref name="col"/>, or EVERY cell in the row when <paramref name="col"/> is null — the shape the
    /// BC look is actually made of (a rule under the whole header row, a rule above a whole totals row), so
    /// the common case is one call rather than one per column. Cosmetic only: no cell, row, grid, span or
    /// binding is touched, which is also why this is the one table operation that does NOT reject
    /// vMerge tables — there is no span arithmetic here to get wrong.
    /// </summary>
    /// <param name="col">
    /// 0-based PHYSICAL cell index within the row (the same index <c>get_layout_info</c> reports), or null
    /// for every cell in the row. Physical rather than grid: a rule is drawn on a cell, and a spanned cell is
    /// one cell however many grid columns it covers.
    /// </param>
    /// <exception cref="ArgumentException">
    /// <paramref name="options"/> selects no edge at all (nothing to do), or its
    /// <see cref="CellBorderOptions.SizeEighthPoints"/> is outside 2–96.
    /// </exception>
    /// <exception cref="NotFoundException">
    /// <paramref name="tableIndex"/>/<paramref name="row"/>/<paramref name="col"/> is out of range, or the
    /// row resolves to no <c>w:tr</c> at all.
    /// </exception>
    public static TableEditResult SetCellBorders(
        WordprocessingDocument doc, LayoutPart part, string? partName, int tableIndex, int row, int? col,
        CellBorderOptions options)
    {
        ArgumentNullException.ThrowIfNull(doc);
        ArgumentNullException.ThrowIfNull(options);

        if (!options.AnyEdge)
        {
            throw new ArgumentException(
                "set_cell_borders needs at least one edge: pass edges as a comma-separated list of "
                + "'top'/'bottom'/'left'/'right' (or 'all').",
                "edges");
        }

        if (options.SizeEighthPoints is < 2 or > 96)
        {
            throw new ArgumentException(
                $"size must be between 2 and 96 eighths of a point (got {options.SizeEighthPoints}); the "
                + $"BC-standard rule is {CellBorderOptions.DefaultSizeEighthPoints} (½ pt).",
                "size");
        }

        var (table, partFile, gridCount) = ResolveTable(doc, part, partName, tableIndex);

        var rowSlots = TableGridNavigator.Rows(table);
        if (row < 0 || row >= rowSlots.Count)
        {
            throw new NotFoundException(
                $"Row index {row} is out of range for table {tableIndex}; it has {rowSlots.Count} row(s).",
                NotFoundTarget.TableCoordinate);
        }

        var innerRow = rowSlots[row].InnerRow
            ?? throw new NotFoundException(
                $"Table {tableIndex}, row {row} resolves to a row-level control with no inner table row to "
                + "operate on.",
                NotFoundTarget.TableCoordinate);

        var cellSlots = TableGridNavigator.Cells(innerRow);
        if (col is { } single && (single < 0 || single >= cellSlots.Count))
        {
            throw new NotFoundException(
                $"Column index {single} is out of range for table {tableIndex}, row {row}; it has "
                + $"{cellSlots.Count} cell(s).",
                NotFoundTarget.TableCoordinate);
        }

        var targets = col is { } only
            ? new[] { cellSlots[only] }
            : cellSlots.ToArray();

        var cellsAffected = 0;
        foreach (var slot in targets)
        {
            if (slot.InnerCell is null)
            {
                continue; // A cell-level control wrapping nothing — nothing to draw a rule on.
            }

            ApplyCellBorders(slot.InnerCell, options);
            cellsAffected++;
        }

        var verb = options.Remove ? "Cleared" : "Drew";
        var where = col is { } c ? $"cell {c} of row {row}" : $"all {cellsAffected} cell(s) of row {row}";
        return new TableEditResult
        {
            Operation = "set_cell_borders",
            Part = partFile,
            TableIndex = tableIndex,
            ColumnIndex = col,
            RowsAffected = 1,
            ColumnCountBefore = gridCount,
            ColumnCountAfter = gridCount,
            Summary = $"{verb} the {DescribeEdges(options)} rule(s) on {where} of table {tableIndex}"
                + (partFile == "document.xml" ? string.Empty : $" in {partFile}")
                + ".",
        };
    }

    private static string DescribeEdges(CellBorderOptions options)
    {
        var edges = new List<string>(4);
        if (options.Top)
        {
            edges.Add("top");
        }

        if (options.Bottom)
        {
            edges.Add("bottom");
        }

        if (options.Left)
        {
            edges.Add("left");
        }

        if (options.Right)
        {
            edges.Add("right");
        }

        return string.Join("/", edges);
    }

    /// <summary>
    /// Applies <paramref name="options"/> to <paramref name="cell"/>'s own <c>w:tcBorders</c>, creating it in
    /// its schema-ordered position when absent and leaving every UNSELECTED edge exactly as it was (so a
    /// header rule added to a cell that already has a left rule keeps both). A selected edge is either drawn
    /// as a single line of the requested thickness or, for <see cref="CellBorderOptions.Remove"/>, set to an
    /// explicit <c>nil</c> — never deleted, since a deleted edge silently reappears in a table whose
    /// <c>w:tblBorders</c> the cell inherits from. Shared with <see cref="SdtFactory.BuildRepeaterTable"/>'s
    /// BC look so a freshly authored header rule and one drawn later are byte-identical.
    /// </summary>
    internal static void ApplyCellBorders(TableCell cell, CellBorderOptions options)
    {
        ArgumentNullException.ThrowIfNull(cell);
        ArgumentNullException.ThrowIfNull(options);

        var tcPr = EnsureTcPr(cell);
        var borders = tcPr.GetFirstChild<TableCellBorders>();
        if (borders is null)
        {
            borders = new TableCellBorders();
            InsertTcPrChildInOrder(tcPr, borders);
        }

        if (options.Top)
        {
            borders.TopBorder = Edge<TopBorder>(options);
        }

        if (options.Bottom)
        {
            borders.BottomBorder = Edge<BottomBorder>(options);
        }

        if (options.Left)
        {
            borders.LeftBorder = Edge<LeftBorder>(options);
        }

        if (options.Right)
        {
            borders.RightBorder = Edge<RightBorder>(options);
        }
    }

    /// <summary>
    /// One border edge in the exact corpus shape — <c>w:val="single" w:sz="4" w:space="0" w:color="auto"</c>
    /// — or an explicit <c>w:val="nil"</c> when clearing (see <see cref="ApplyCellBorders"/>).
    /// </summary>
    private static TEdge Edge<TEdge>(CellBorderOptions options)
        where TEdge : BorderType, new() =>
        new()
        {
            Val = options.Remove ? BorderValues.Nil : BorderValues.Single,
            Size = (uint)options.SizeEighthPoints,
            Space = 0,
            Color = "auto",
        };

    /// <summary>
    /// Inserts ONE new grid column into the <paramref name="tableIndex"/>-th table at grid position
    /// <paramref name="atColumn"/> (the far-right edge when omitted): one new <c>w:gridCol</c> plus one new
    /// cell per row — a header cell in a header row, the bound control cell in the data rows (the rows of the
    /// repeater <see cref="InsertColumnOptions.DataPath"/> belongs to, or, in a table with no repeater, every
    /// non-header row), and an empty cell everywhere else. See <see cref="PlanInsertColumnRows"/> for which
    /// rows are which.
    /// </summary>
    /// <remarks>
    /// SPANNED CELLS AT AN INTERIOR POSITION. Real BC tables are pervasively <c>gridSpan</c>-spanned, so an
    /// interior position can fall in the MIDDLE of a cell rather than on a boundary. Two different rules
    /// apply, by what the row is for:
    /// <list type="bullet">
    /// <item>A row with no content of its own there (a spacer, or a totals block's leading empty run) has its
    /// spanning cell WIDENED by one column instead of gaining a cell — every other cell keeps its position
    /// and width, so a right-anchored summary stays anchored to the table's right edge.</item>
    /// <item>A row that must RECEIVE the new column's content (a header row, or the data row for a
    /// field/label column) is REFUSED, naming the offending cell: widening its spanned cell would silently
    /// drop the content the caller asked for, and splitting it would rewrite a layout decision the caller
    /// never mentioned. <c>split_cells</c> makes the boundary first if that is what was meant.</item>
    /// </list>
    /// </remarks>
    /// <exception cref="ArgumentException">
    /// <paramref name="atColumn"/> is out of range; an interior <paramref name="atColumn"/> falls inside a
    /// spanned cell of a row that must receive content (see remarks); <paramref name="options"/> is
    /// inconsistent (e.g. Field/Label without a DataPath); DataPath belongs to none of the table's repeaters,
    /// or the table has no row to bind it into; a header argument is passed for a table with no header row;
    /// the table uses vMerge; or (propagated from <see cref="SdtFactory"/>) DataPath or HeaderLabelPath does
    /// not resolve to a leaf column of the right label-shape.
    /// </exception>
    public static TableEditResult InsertColumn(
        WordprocessingDocument doc, LayoutPart part, string? partName, int tableIndex, int? atColumn, InsertColumnOptions options)
    {
        ArgumentNullException.ThrowIfNull(doc);
        ArgumentNullException.ThrowIfNull(options);

        var (table, partFile, gridCount) = ResolveTable(doc, part, partName, tableIndex);

        // atColumn omitted (null) means "append at the far-right edge".
        var target = atColumn ?? gridCount;
        if (target < 0 || target > gridCount)
        {
            throw new ArgumentException(
                $"atColumn {target} is out of range for table {tableIndex}; it has {gridCount} grid "
                + $"column(s), so atColumn must be between 0 and {gridCount} ({gridCount} = append at the far-right edge).",
                nameof(atColumn));
        }

        RejectUnsupportedShape(table, "insert_column");

        var mode = options.Mode;
        if (mode is InsertColumnMode.Field or InsertColumnMode.Label && string.IsNullOrWhiteSpace(options.DataPath))
        {
            throw new ArgumentException(
                $"insert_column mode '{mode}' requires a dataPath (the dataset path to bind the new column's "
                + "data cell to, e.g. '/Header/Line/Discount_Line').",
                "dataPath");
        }

        var width = options.Width ?? MeanGridWidth(table);

        // The schema is needed for a bound data cell AND for a bound header label - plainText mode with a
        // headerLabelPath used to skip loading it, which silently turned the label into static text and
        // never checked the path at all (GitHub issue #82).
        var schema = mode != InsertColumnMode.PlainText || options.HeaderLabelPath is not null
            ? SchemaProvider.FromLayout(doc)
            : null;
        ValidateInsertColumnPaths(schema, options);

        var nextId = LayoutEditor.MakeIdGenerator(doc);
        var plans = PlanInsertColumnRows(table, tableIndex, options);

        // Refuse BEFORE mutating anything: an interior position that falls inside a spanned cell of a row
        // that must carry the new column's content has no honest answer (see this method's remarks).
        RejectStraddledContentRows(plans, target, tableIndex);

        var rowsAffected = 0;
        var rowsToReslice = new List<TableRow>();
        foreach (var plan in plans)
        {
            Paragraph paragraph;
            if (plan.IsHeaderRow)
            {
                paragraph = BuildHeaderParagraph(schema, options, nextId);
            }
            else if (plan.ReceivesContent)
            {
                var control = mode == InsertColumnMode.Label
                    ? (OpenXmlElement)SdtFactory.BuildLabel(schema!, options.DataPath!, id: nextId())
                    : SdtFactory.BuildField(schema!, options.DataPath!, id: nextId());
                paragraph = new Paragraph(control);
            }
            else
            {
                paragraph = new Paragraph();
            }

            var rowCells = TableGridNavigator.Cells(plan.Row);
            var isFillerRow = !plan.ReceivesContent;
            var coverage = TableGridNavigator.Coverage(plan.Row);

            // A row whose skipped run covers the new grid position grows that run instead of gaining a cell.
            // Checked before the totals-block and straddle paths below: on such a row those would both be
            // reasoning about cells that do not extend to the insertion point.
            if (TryGrowSkippedRegion(plan.Row, coverage, target, plan.ReceivesContent))
            {
                rowsToReslice.Add(plan.Row);
                rowsAffected++;
                continue;
            }


            // A filler row shaped like a BC totals/summary block — leading EMPTY cell(s), content hugging
            // the right edge (e.g. "Total Incl VAT | 822.97" under a lines table) — must KEEP that content
            // at the table's right edge: appending an empty cell at the row's end would push the summary
            // one column in from the edge AND extend the block's horizontal rules (cloned borders) under
            // the new column. Instead of adding any cell, the row's TRAILING content cell is widened
            // (gridSpan +1) across the new column — every cell keeps its position and width, only the
            // last one stretches, so the block looks exactly as before with its right-aligned content
            // still on the table's right edge. Appends only: an interior position lands inside the block
            // rather than past it, where the straddle rule below already keeps the layout intact.
            var appending = target == gridCount;
            if (appending && isFillerRow && rowCells.Count > 0 && IsEmptyCell(rowCells[0]) && !IsEmptyCell(rowCells[^1]))
            {
                SetCellGridSpan(rowCells[^1].InnerCell, rowCells[^1].GridSpan + 1);
                rowsToReslice.Add(plan.Row);
                rowsAffected++;
                continue;
            }

            var slot = LocateInsertion(rowCells, target, coverage.Before);
            if (slot.StraddlingCell is { } straddled)
            {
                // A filler row whose cell spans across the insertion point simply gets wider — no cell is
                // added, so nothing shifts and nothing is lost (content rows never reach here: they were
                // refused above).
                SetCellGridSpan(straddled.InnerCell, straddled.GridSpan + 1);
                rowsToReslice.Add(plan.Row);
                rowsAffected++;
                continue;
            }

            // The new cell copies the look of an existing neighbor — preferring the one on its LEFT, the
            // same choice the append path has always made (there, the row's last cell). Without this, the
            // new column visibly doesn't belong to the table: no header underline (per-cell w:tcBorders in
            // real BC layouts), top-aligned text next to bottom-aligned neighbors, left-aligned numbers in
            // an otherwise right-aligned column.
            var neighbor = NeighborCell(rowCells, slot.CellIndex);
            var newCell = BuildCellLike(neighbor, paragraph, width);
            if (slot.CellIndex >= rowCells.Count)
            {
                plan.Row.AppendChild(newCell);
            }
            else
            {
                plan.Row.InsertBefore(newCell, rowCells[slot.CellIndex].CellChild);
            }

            rowsAffected++;
        }

        var newGridColumn = new GridColumn { Width = width.ToString(CultureInfo.InvariantCulture) };
        var gridColumns = GridColumns(table);
        if (target >= gridColumns.Count)
        {
            GridElement(table).AppendChild(newGridColumn);
        }
        else
        {
            GridElement(table).InsertBefore(newGridColumn, gridColumns[target]);
        }

        // A widened cell now also covers the new column, so its w:tcW must grow by that column's width —
        // re-slice the affected rows against the final grid (skipped when any declared grid width doesn't
        // parse; there is nothing reliable to slice by).
        if (rowsToReslice.Count > 0 && TryGetGridWidths(table, out var finalWidths))
        {
            foreach (var widened in rowsToReslice)
            {
                ResliceRowWidths(widened, finalWidths);
            }
        }

        SyncTableWidthToGrid(table);

        var placement = target == gridCount
            ? $"Appended a new {DescribeMode(options)} column (grid index {target})"
            : $"Inserted a new {DescribeMode(options)} column at grid index {target}";

        return new TableEditResult
        {
            Operation = "insert_column",
            Part = partFile,
            TableIndex = tableIndex,
            ColumnIndex = target, // the new column's 0-based grid index
            RowsAffected = rowsAffected,
            ColumnCountBefore = gridCount,
            ColumnCountAfter = gridCount + 1,
            Summary = $"{placement} to table {tableIndex}"
                + (partFile == "document.xml" ? string.Empty : $" in {partFile}")
                + $", adding a cell to {rowsAffected} row(s).",
        };
    }

    /// <summary>One row's role in an <see cref="InsertColumn"/> call, decided once and reused by the pre-flight check and the edit itself.</summary>
    private readonly record struct RowInsertPlan(
        int RowIndex, TableRow Row, bool IsHeaderRow, bool IsDataRow, bool ReceivesContent);

    /// <summary>
    /// Checks every dataset path an <see cref="InsertColumn"/> call will bind BEFORE any row is touched:
    /// <see cref="InsertColumnOptions.DataPath"/> in field/label mode, and
    /// <see cref="InsertColumnOptions.HeaderLabelPath"/> in EVERY mode. Building the control is the
    /// validation (<see cref="SdtFactory"/> owns the leaf-column and label-shape rules); the throwaway id
    /// keeps the document's id generator untouched.
    /// </summary>
    /// <remarks>
    /// Up front rather than where the controls are built so that a bad path is reported as itself, not as a
    /// side effect of row planning (a misspelled dataPath would otherwise surface as "no repeater of this
    /// table contains it"), and so that a headerLabelPath is checked even on a table that turns out to have
    /// no header row. A headerLabelPath failure is re-raised under its own argument name: the caller has two
    /// paths in one call and must be told which one to fix.
    /// </remarks>
    private static void ValidateInsertColumnPaths(DatasetTree? schema, InsertColumnOptions options)
    {
        if (schema is null)
        {
            return;
        }

        switch (options.Mode)
        {
            case InsertColumnMode.Field:
                _ = SdtFactory.BuildField(schema, options.DataPath!, id: 0);
                break;
            case InsertColumnMode.Label:
                _ = SdtFactory.BuildLabel(schema, options.DataPath!, id: 0);
                break;
        }

        if (options.HeaderLabelPath is { } headerLabelPath)
        {
            try
            {
                _ = SdtFactory.BuildLabel(schema, headerLabelPath, id: 0);
            }
            catch (ArgumentException ex) when (ex.ParamName == "datasetPath")
            {
                throw new ArgumentException($"headerLabelPath: {ex.Message}", "headerLabelPath", ex);
            }
        }
    }

    /// <summary>
    /// Decides each row's role for <see cref="InsertColumn"/>: header cell, bound data cell, or filler (an
    /// empty cell, or no cell at all where the totals-block / straddle / skipped-run rules apply).
    /// </summary>
    /// <remarks>
    /// <para>
    /// HEADER ROWS. Rows marked <c>w:tblHeader</c> when the table has any; otherwise, in a table with a
    /// repeater, its first row when that row is not a repeater row (BC line tables mostly carry no
    /// <c>w:tblHeader</c> — StandardSalesQuote's is the exception — and always start with the caption row).
    /// A PLAIN table has no such convention to lean on: its header rows are its CAPTION rows — rows whose
    /// bound controls are all labels — which is how every caption-over-value block in the corpus is built
    /// (StandardSalesQuote's "Your Reference / Valid To / ..." table, the purchase order's two stacked
    /// blocks, the footers' bank-detail pairs). "Row 0 is the header" was the alternative, and the corpus
    /// rules it out: row 0 of the address blocks and the caption-beside-value tables is a FIELD row, so it
    /// would have put a caption above a customer address. A stacked second caption row gets the header too,
    /// so each block keeps a caption over its value.
    /// </para>
    /// <para>
    /// DATA ROWS. In a repeater table, ONLY the rows of the repeater the data path belongs to (see
    /// <see cref="ResolveDataRepeater"/>): the stock 1304/1306 lines tables also hold the
    /// <c>/Header/ReportTotalsLine</c> repeater, which used to receive the <c>/Header/Line</c> field as
    /// well — rendering Line 1's value beside every totals line, and skipping the totals-block widening that
    /// keeps those amounts on the right edge (issue #82). Every other repeater's rows are filler now, so
    /// that rule applies to them. In a plain table, every row that is not a header row, as before.
    /// </para>
    /// <para>
    /// UNUSABLE ARGUMENTS ARE REFUSED, not dropped: an explicit headerText/headerLabelPath on a table with no
    /// header row (it used to vanish, returning <c>ok:true</c>), and field/label mode on a table with no row
    /// to hold the field.
    /// </para>
    /// <para>
    /// Only the FIRST <c>w:tr</c> of a repeater's item is visible here (<see cref="TableGridNavigator"/>
    /// resolves a row-level control to one inner row; GitHub issue #80), so a multi-row item's later rows are
    /// not planned at all — unchanged by this method.
    /// </para>
    /// </remarks>
    private static List<RowInsertPlan> PlanInsertColumnRows(Table table, int tableIndex, InsertColumnOptions options)
    {
        var mode = options.Mode;
        var rowSlots = TableGridNavigator.Rows(table);
        var repeaterPaths = rowSlots.Select(r => r.InnerRow is null ? null : RepeaterDataItemPath(r)).ToList();
        var hasRepeater = repeaterPaths.Any(p => p is not null);
        var anyTblHeader = rowSlots.Any(r => HasTblHeader(r.InnerRow));
        var dataRepeater = hasRepeater && mode != InsertColumnMode.PlainText
            ? ResolveDataRepeater(repeaterPaths, options.DataPath!, tableIndex)
            : null;

        var plans = new List<RowInsertPlan>(rowSlots.Count);
        for (var idx = 0; idx < rowSlots.Count; idx++)
        {
            if (rowSlots[idx].InnerRow is not { } innerRow)
            {
                continue;
            }

            var isHeaderRow = anyTblHeader
                ? HasTblHeader(innerRow)
                : hasRepeater
                    ? idx == 0 && repeaterPaths[idx] is null
                    : IsCaptionRow(innerRow);
            var isDataRow = !isHeaderRow
                && (hasRepeater
                    ? dataRepeater is not null && string.Equals(repeaterPaths[idx], dataRepeater, StringComparison.Ordinal)
                    : mode != InsertColumnMode.PlainText);
            var receivesContent = isHeaderRow || (isDataRow && mode != InsertColumnMode.PlainText);

            plans.Add(new RowInsertPlan(idx, innerRow, isHeaderRow, isDataRow, receivesContent));
        }

        if ((options.HeaderText is not null || options.HeaderLabelPath is not null) && !plans.Any(p => p.IsHeaderRow))
        {
            var argument = options.HeaderLabelPath is not null ? "headerLabelPath" : "headerText";
            throw new ArgumentException(
                $"Table {tableIndex} has no header row to put {argument} in: no row is marked w:tblHeader, "
                + (hasRepeater
                    ? "and its first row is a repeater row"
                    : "and no row is a caption row (one whose bound controls are all labels)")
                + $". Omit {argument}, or add the caption after the insert with set_cell_text/insert_label at "
                + "the cell you mean.",
                argument);
        }

        if (mode != InsertColumnMode.PlainText && !plans.Any(p => p.IsDataRow))
        {
            throw new ArgumentException(
                $"Table {tableIndex} has no row to bind '{options.DataPath}' into: every row is a header "
                + "(caption) row. Add the column in plainText mode and bind the cell you want with "
                + "insert_field/insert_label.",
                "dataPath");
        }

        return plans;
    }

    /// <summary>
    /// The repeater whose rows receive a field/label column bound to <paramref name="dataPath"/>: the
    /// table repeater whose data item is the NEAREST ancestor of the path's own data item (for
    /// <c>/Header/Line/ItemReferenceNo_Line</c> in a table repeating <c>/Header/Line</c> and
    /// <c>/Header/ReportTotalsLine</c>, that is <c>/Header/Line</c>; with nested repeaters, the deepest one
    /// that still contains the path).
    /// </summary>
    /// <remarks>
    /// A path OUTSIDE every repeater of the table (a parent/Header field shown on each line, which the tool
    /// contract allows) has no nearest repeater. It is placed only when the table repeats a single data
    /// item nested under the path's own — then there is one candidate row and no choice to make. With two
    /// or more (the stock lines tables' Line + ReportTotalsLine), or when the path is unrelated to every
    /// repeater, the call is refused naming the table's repeaters rather than guessing (ADR-0003).
    /// </remarks>
    private static string ResolveDataRepeater(IReadOnlyList<string?> repeaterPaths, string dataPath, int tableIndex)
    {
        var pathSegments = dataPath.Split('/', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        var dataItem = "/" + string.Join("/", pathSegments.Take(pathSegments.Length - 1));
        var repeaters = repeaterPaths.OfType<string>().Distinct(StringComparer.Ordinal).ToList();

        var enclosing = repeaters
            .Where(r => IsSameOrDescendant(dataItem, r))
            .OrderByDescending(r => r.Length)
            .FirstOrDefault();
        if (enclosing is not null)
        {
            return enclosing;
        }

        if (repeaters.Count == 1 && IsSameOrDescendant(repeaters[0], dataItem))
        {
            return repeaters[0];
        }

        throw new ArgumentException(
            $"dataPath '{dataPath}' does not belong to a repeater of table {tableIndex}, which repeats "
            + string.Join(", ", repeaters.Select(r => $"'{r}'"))
            + ". insert_column binds the field only into the rows of the repeater whose data item contains "
            + "dataPath (a parent data item's field only when the table has a single repeater under it, so "
            + "the target row is unambiguous). Pick a column of one of those data items, or bind a specific "
            + "cell with insert_field (locationType 'tableCell').",
            nameof(dataPath));

        static bool IsSameOrDescendant(string path, string ancestor) =>
            string.Equals(path, ancestor, StringComparison.Ordinal)
            || path.StartsWith(ancestor + "/", StringComparison.Ordinal);
    }

    /// <summary>
    /// The dataset path (<c>/Header/Line</c>) of the data item a row repeats over: the binding of the
    /// INNERMOST repeating section in the row's wrapper chain (the one closest to the <c>w:tr</c>, i.e. the
    /// item the row is cloned for). Null for a bare row, and for a row-level control that is not a
    /// repeater (or carries no binding) — such a row is planned like any other plain row.
    /// </summary>
    private static string? RepeaterDataItemPath(TableGridNavigator.RowSlot row)
    {
        for (var i = row.Wrappers.Count - 1; i >= 0; i--)
        {
            var wrapper = row.Wrappers[i];
            if (!SdtInspector.IsRepeater(wrapper))
            {
                continue;
            }

            var binding = wrapper.GetFirstChild<SdtProperties>() is { } pr ? SdtInspector.FindRepeaterBinding(pr) : null;
            var segments = BindingXPath.Segments(binding is null ? null : SdtInspector.Attr(binding, "xpath", OoxmlNames.W));
            return segments.Count <= 1 ? null : "/" + string.Join("/", segments.Skip(1));
        }

        return null;
    }

    /// <summary>
    /// True for a plain table's CAPTION row: it holds at least one bound label control and no other bound
    /// control (field, picture, repeater). Literal text does not count — a hand-typed row is
    /// indistinguishable from a row of values, and BC's own captions are label-bound.
    /// </summary>
    private static bool IsCaptionRow(TableRow innerRow)
    {
        var kinds = innerRow.Descendants<SdtElement>().Select(SdtInspector.ClassifyControlKind).ToList();
        return kinds.Contains(ControlKind.Label)
            && !kinds.Any(k => k is ControlKind.Field or ControlKind.Picture or ControlKind.Repeater);
    }

    /// <summary>
    /// Where a new cell covering grid column <paramref name="target"/> goes in one row:
    /// <see cref="InsertionSlot.CellIndex"/> is the physical cell index to insert BEFORE (equal to the cell
    /// count when the new column lands at the row's end), unless an existing cell spans ACROSS that grid
    /// position — then <see cref="InsertionSlot.StraddlingCell"/> is that cell and no boundary exists.
    /// </summary>
    private readonly record struct InsertionSlot(int CellIndex, TableGridNavigator.CellSlot? StraddlingCell);

    private static InsertionSlot LocateInsertion(
        IReadOnlyList<TableGridNavigator.CellSlot> cells, int target, int gridBefore = 0)
    {
        // Physical cells start at grid column `gridBefore`, so a target left of that reaches none of them;
        // the nearest honest boundary is the row's own start.
        if (target < gridBefore)
        {
            return new InsertionSlot(0, null);
        }

        var cursor = gridBefore;
        for (var i = 0; i < cells.Count; i++)
        {
            if (cursor == target)
            {
                return new InsertionSlot(i, null);
            }

            if (cursor < target && cursor + cells[i].GridSpan > target)
            {
                return new InsertionSlot(-1, cells[i]);
            }

            cursor += cells[i].GridSpan;
        }

        // cursor == target (the row ends exactly at the insertion point) or, for a pre-existing ragged row
        // whose coverage falls short of the grid, the closest honest answer: append at the row's end.
        return new InsertionSlot(cells.Count, null);
    }

    /// <summary>
    /// Grows a FILLER row's leading or trailing skipped run by one when the new grid column lands inside it,
    /// and reports whether it did. Such a row reaches that grid position with no cell at all, so adding one
    /// would put a cell where the row deliberately has none — growing the run keeps the row looking exactly as
    /// it did while still accounting for the extra grid column.
    /// </summary>
    /// <remarks>
    /// Deliberately NOT applied to a row that must receive the new column's content. On those, growing a
    /// skipped run would silently discard the field/label the caller asked for — the same failure
    /// <see cref="RejectStraddledContentRows"/> refuses to commit. A content row instead gets a real cell at
    /// the nearest boundary <see cref="LocateInsertion"/> finds (its own first or last cell position when the
    /// target is outside its covered span), leaving its skipped run untouched: on the stock
    /// <c>StandardSalesInvoiceVatSpec.docx</c> lines table — three rows of ten cells plus one skipped column,
    /// one row of eleven — appending a column therefore puts the new cell immediately after each content row's
    /// last cell and shifts the blank filler column right, which is what appending a column to that table
    /// looks like when a person does it in Word.
    /// </remarks>
    private static bool TryGrowSkippedRegion(
        TableRow innerRow, TableGridNavigator.RowGridCoverage coverage, int target, bool receivesContent)
    {
        if (receivesContent)
        {
            return false;
        }

        if (coverage.Before > 0 && target < coverage.Before)
        {
            TableGridNavigator.SetGridEdges(innerRow, coverage.Before + 1, coverage.After);
            return true;
        }

        if (coverage.After > 0 && target >= coverage.ContentEnd)
        {
            TableGridNavigator.SetGridEdges(innerRow, coverage.Before, coverage.After + 1);
            return true;
        }

        return false;
    }

    /// <summary>The cell whose look a new cell at <paramref name="cellIndex"/> should copy: its left neighbor, else its right.</summary>
    private static TableCell? NeighborCell(IReadOnlyList<TableGridNavigator.CellSlot> cells, int cellIndex)
    {
        if (cells.Count == 0)
        {
            return null;
        }

        return cellIndex > 0 ? cells[Math.Min(cellIndex, cells.Count) - 1].InnerCell : cells[0].InnerCell;
    }

    /// <summary>
    /// Refuses an interior insertion whose grid position falls INSIDE a spanned cell of a row that must
    /// receive the new column's content — see <see cref="InsertColumn"/>'s remarks for why neither widening
    /// nor splitting is an honest answer there. Runs before any mutation, so a refused call leaves the
    /// document exactly as it was.
    /// </summary>
    private static void RejectStraddledContentRows(IReadOnlyList<RowInsertPlan> plans, int target, int tableIndex)
    {
        foreach (var plan in plans.Where(p => p.ReceivesContent))
        {
            var cells = TableGridNavigator.Cells(plan.Row);
            var coverage = TableGridNavigator.Coverage(plan.Row);
            var slot = LocateInsertion(cells, target, coverage.Before);
            if (slot.StraddlingCell is not { } straddled)
            {
                continue;
            }

            var cellIndex = cells.ToList().FindIndex(c => ReferenceEquals(c.CellChild, straddled.CellChild));
            var start = coverage.Before + cells.Take(cellIndex).Sum(c => c.GridSpan);
            throw new ArgumentException(
                $"atColumn {target} falls INSIDE a spanned cell of table {tableIndex}, row {plan.RowIndex}: "
                + $"cell {cellIndex} spans grid columns {start}..{start + straddled.GridSpan - 1}, and that row "
                + "must carry the new column's content (it is a header row or the bound data row). Widening "
                + "that cell would silently drop the content you asked for, and splitting it would change a "
                + "layout decision you did not ask about. Split that cell first with split_cells, or pick an "
                + "atColumn that lands on a cell boundary in every content row.",
                "atColumn");
        }
    }

    /// <summary>
    /// Inserts ONE static (non-repeating) <c>w:tr</c> into the <paramref name="tableIndex"/>-th table at
    /// row-slot position <paramref name="atRow"/> (0-based, the SAME numbering <c>get_layout_info</c>'s
    /// <c>tables[]</c> uses, where a repeater's whole repeating section counts as one row; null appends
    /// after the last row) — the stock BC shape for the totals block INSIDE the line-items table (GitHub
    /// issue #28): every stock document layout ends its lines table with right-anchored trailing rows
    /// (leading empty cells, <c>gridSpan</c>-merged content cells, bound totals fields — corpus:
    /// <c>StandardSalesQuote</c>/<c>StandardPurchaseOrder</c>/<c>StandardSalesInvoiceVatSpec</c>) rather
    /// than the separate totals table that used to be the only buildable route. Cell construction is
    /// <see cref="SdtFactory.BuildStaticRow"/>'s (each <see cref="RepeaterRowCell.Columns"/> entry a FULL
    /// dataset path; label-shaped paths become labels); the spans must cover the grid exactly, which the
    /// grid-consistency guard also enforces mechanically before anything reaches disk. The optional
    /// <paramref name="format"/> styles every bound cell's runs (the corpus grand-total rows are bold) via
    /// the same path <c>insert_field</c> uses; a literal caption or a rule above the row are follow-up
    /// <c>set_cell_text</c>/<c>set_cell_borders</c> calls on the reported row.
    /// </summary>
    /// <remarks>
    /// The row is inserted as a DIRECT child of the <c>w:tbl</c>, never inside a repeating section — that
    /// is what makes it render exactly once. It is deliberately NOT wrapped in any <c>w:sdt</c>: the
    /// corpus totals rows' own cell-level sdt wrappers are a binding vehicle, and this factory's inline
    /// controls (the shape every tool-built bound cell already uses, BC-verified in the sandbox rounds)
    /// carry the same bindings without them. Rows INSIDE a repeater's item (per-group subtotals) are the
    /// sibling operation <see cref="LayoutEditor.InsertSubtotalRow"/>.
    /// </remarks>
    /// <exception cref="ArgumentException">
    /// <paramref name="cells"/> is empty or its spans do not sum to the grid column count;
    /// <paramref name="atRow"/> is out of range; <paramref name="format"/> sets an alignment (per-cell
    /// alignments belong on <paramref name="cells"/>); the table uses vMerge; or (propagated from
    /// <see cref="SdtFactory.BuildStaticRow"/>) a path does not resolve to a leaf column.
    /// </exception>
    /// <exception cref="NotFoundException"><paramref name="tableIndex"/> is out of range for the part.</exception>
    public static TableEditResult InsertStaticRow(
        WordprocessingDocument doc,
        LayoutPart part,
        string? partName,
        int tableIndex,
        int? atRow,
        IReadOnlyList<RepeaterRowCell> cells,
        CellTextFormat? format = null)
    {
        ArgumentNullException.ThrowIfNull(doc);
        ArgumentNullException.ThrowIfNull(cells);

        if (cells.Count == 0)
        {
            throw new ArgumentException("At least one cell spec is required (cells was empty).", nameof(cells));
        }

        if (format?.Alignment is not null)
        {
            // Checked HERE rather than left to ApplyControlRunFormat below, which only runs for BOUND
            // cells: an all-spacer row (a legitimate call — the corpus spacer rows) would otherwise
            // silently swallow the unsupported knob instead of refusing it.
            throw new ArgumentException(
                "alignment is not supported as a row-level format on insert_table_row; pass per-cell "
                + "alignments alongside the cell specs instead (one 'left'/'center'/'right' or '-' per cell).",
                nameof(format));
        }

        var (table, partFile, gridCount) = ResolveTable(doc, part, partName, tableIndex);
        RejectUnsupportedShape(table, "insert_table_row");

        var spanSum = cells.Sum(c => c.Span);
        if (spanSum != gridCount)
        {
            throw new ArgumentException(
                $"The cells' spans sum to {spanSum} but table {tableIndex} has {gridCount} grid column(s); "
                + "a static row must cover the grid exactly (use a spacer cell '-' with a span for the "
                + "unused width, e.g. '4:-').",
                nameof(cells));
        }

        var rowSlots = TableGridNavigator.Rows(table);
        var target = atRow ?? rowSlots.Count;
        if (target < 0 || target > rowSlots.Count)
        {
            throw new ArgumentException(
                $"atRow {target} is out of range for table {tableIndex}; it has {rowSlots.Count} row(s), so "
                + $"atRow must be between 0 and {rowSlots.Count} ({rowSlots.Count} = append after the last "
                + "row, where a totals row goes).",
                nameof(atRow));
        }

        // ResolveTable already guarantees a w:tblGrid exists (gridCount comes from it).
        var gridColumns = table.GetFirstChild<TableGrid>()!.Elements<GridColumn>().ToList();
        var widths = LayoutEditor.ComputeCellWidthsFromGrid(gridColumns, cells);

        var schema = SchemaProvider.FromLayout(doc);
        var nextId = LayoutEditor.MakeIdGenerator(doc);
        var row = SdtFactory.BuildStaticRow(schema, cells, widths, nextId);

        var boundControlIds = new List<int>();
        foreach (var sdt in row.Descendants<SdtRun>())
        {
            LayoutEditor.ApplyControlRunFormat(sdt, format, "insert_table_row");
            if (SdtInspector.ReadControlId(sdt) is { } id)
            {
                boundControlIds.Add(id);
            }
        }

        if (target == rowSlots.Count)
        {
            // Rows are the last thing CT_Tbl's content model allows, so appending lands after the
            // existing rows regardless of whether any exist.
            table.AppendChild(row);
        }
        else
        {
            table.InsertBefore(row, rowSlots[target].RowChild);
        }

        var boundSummary = boundControlIds.Count == 0
            ? "no bound cells (a spacer row)"
            : $"{boundControlIds.Count} bound control(s) (id(s) {string.Join(", ", boundControlIds)})";

        return new TableEditResult
        {
            Operation = "insert_table_row",
            Part = partFile,
            TableIndex = tableIndex,
            ColumnIndex = null,
            RowsAffected = 1,
            ColumnCountBefore = gridCount,
            ColumnCountAfter = gridCount,
            Summary = $"Inserted a static (non-repeating) row at row index {target} of table {tableIndex} "
                + $"with {cells.Count} cell(s) and {boundSummary}. It renders exactly once - it is a direct "
                + "table row, not part of any repeating section. Follow up with set_cell_borders (row "
                + $"{target}) for the rule a BC totals row carries, or set_cell_text for a literal caption.",
        };
    }

    /// <summary>
    /// Removes the <paramref name="gridColumn"/>-th grid column of the <paramref name="tableIndex"/>-th
    /// table: the matching <c>w:gridCol</c>, and in every row the physical cell covering that grid column —
    /// deleted outright when it spans only that column (dropping a bound cell too, unlike
    /// <see cref="LayoutEditor.RemoveControl"/>, which preserves the cell), or its <c>w:gridSpan</c>
    /// decremented when it spans more. Repeater/cell wrappers are preserved (only the inner cells/rows are
    /// touched). Every row's coverage stays equal to the shrunk grid, verified by the consistency guard.
    /// </summary>
    /// <exception cref="ArgumentException">
    /// <paramref name="gridColumn"/> is out of range, it is the last remaining column, it is some row's only
    /// cell, no row of the table matches its grid (see <see cref="RejectUnremovableColumn"/>), or the table
    /// uses vMerge.
    /// </exception>
    /// <exception cref="NotFoundException"><paramref name="tableIndex"/> is out of range for the part.</exception>
    public static TableEditResult RemoveColumn(
        WordprocessingDocument doc, LayoutPart part, string? partName, int tableIndex, int gridColumn)
    {
        ArgumentNullException.ThrowIfNull(doc);

        var (table, partFile, gridCount) = ResolveTable(doc, part, partName, tableIndex);

        if (gridColumn < 0 || gridColumn >= gridCount)
        {
            throw new ArgumentException(
                $"gridColumn {gridColumn} is out of range for table {tableIndex}; it has {gridCount} grid "
                + $"column(s) (valid 0..{gridCount - 1}).",
                nameof(gridColumn));
        }

        if (gridCount <= 1)
        {
            throw new ArgumentException(
                $"Cannot remove the last remaining column of table {tableIndex}; a table must keep at least "
                + "one column (removing it would empty every row). Remove the whole table by hand instead if "
                + "that is the intent.",
                nameof(gridColumn));
        }

        RejectUnsupportedShape(table, "remove_column");

        // Refuse BEFORE mutating anything, so a refused call leaves the document exactly as it was.
        RejectUnremovableColumn(table, tableIndex, gridColumn, gridCount);

        // BC line tables span the full content width; silently shrinking the table by the removed
        // column's width is the kind of visual regression a caller only catches in a preview. Capture
        // the pre-removal widths so the removed width can be redistributed proportionally afterwards
        // (skipped only when the grid's declared widths don't all parse — then there is no reliable
        // basis to scale from).
        var allWidthsKnown = TryGetGridWidths(table, out var widthsBefore);

        var rowsAffected = 0;
        foreach (var row in TableGridNavigator.Rows(table))
        {
            if (row.InnerRow is null)
            {
                continue;
            }

            var cells = TableGridNavigator.Cells(row.InnerRow);
            var coverage = TableGridNavigator.Coverage(row.InnerRow);
            if (coverage.Total != gridCount)
            {
                continue; // A pre-existing malformed row: do not guess how to shrink it.
            }

            // The removed grid column may fall in a row's SKIPPED region rather than on one of its cells (a
            // gridBefore/gridAfter row covers fewer physical cells than the grid has columns). Then there is
            // no cell to delete or narrow: the skipped run itself shrinks by one, which keeps this row's
            // coverage equal to the newly shrunk grid exactly as a cell removal does for the other rows.
            if (gridColumn < coverage.Before)
            {
                TableGridNavigator.SetGridEdges(row.InnerRow, coverage.Before - 1, coverage.After);
                rowsAffected++;
                continue;
            }

            if (gridColumn >= coverage.ContentEnd)
            {
                TableGridNavigator.SetGridEdges(row.InnerRow, coverage.Before, coverage.After - 1);
                rowsAffected++;
                continue;
            }

            var cursor = coverage.Before;
            foreach (var cell in cells)
            {
                if (gridColumn >= cursor && gridColumn < cursor + cell.GridSpan)
                {
                    if (cell.GridSpan <= 1)
                    {
                        cell.CellChild.Remove();
                    }
                    else
                    {
                        SetCellGridSpan(cell.InnerCell, cell.GridSpan - 1);
                    }

                    rowsAffected++;
                    break;
                }

                cursor += cell.GridSpan;
            }
        }

        GridColumns(table)[gridColumn].Remove();

        var widthNote = string.Empty;
        if (allWidthsKnown)
        {
            var removedWidth = widthsBefore[gridColumn];
            var newWidths = RedistributeWidth(widthsBefore, gridColumn);
            ApplyGridWidths(table, newWidths);
            widthNote = $"; its {removedWidth} twips were redistributed proportionally across the remaining "
                + $"columns (table width kept at {widthsBefore.Sum()} twips - use set_column_widths for a "
                + "different distribution)";
        }

        SyncTableWidthToGrid(table);

        return new TableEditResult
        {
            Operation = "remove_column",
            Part = partFile,
            TableIndex = tableIndex,
            ColumnIndex = gridColumn,
            RowsAffected = rowsAffected,
            ColumnCountBefore = gridCount,
            ColumnCountAfter = gridCount - 1,
            Summary = $"Removed grid column {gridColumn} from table {tableIndex}"
                + (partFile == "document.xml" ? string.Empty : $" in {partFile}")
                + $", updating {rowsAffected} row(s){widthNote}.",
        };
    }

    /// <summary>
    /// The two shapes <see cref="RemoveColumn"/> must refuse rather than "succeed" on, checked up front.
    /// </summary>
    /// <remarks>
    /// <list type="bullet">
    /// <item>A TABLE NO ROW OF WHICH MATCHES ITS GRID. Rows whose coverage differs from the grid are skipped
    /// (there is no honest way to shrink a row that is already wrong), which is right for one pre-existing
    /// ragged row in an otherwise sound table. When EVERY row is skipped, though, the grid itself is the
    /// outlier (a hand-edited <c>w:tblGrid</c> shorter or longer than the rows), and dropping a
    /// <c>w:gridCol</c> while touching no row only widens the mismatch — the call used to report
    /// <c>ok:true</c> and "updating 0 row(s)" there (GitHub issue #78).</item>
    /// <item>A ROW WHOSE ONLY CELL IS THE ONE TO DELETE. A row covering most of the grid with
    /// <c>w:gridBefore</c>/<c>w:gridAfter</c> can be down to a single span-1 cell; deleting it leaves a
    /// <c>w:tr</c> with no <c>w:tc</c>, which the grid arithmetic still accepts (the skipped run covers
    /// the shrunk grid) and <see cref="DocumentFormat.OpenXml.Validation.OpenXmlValidator"/> accepts too —
    /// but Word cannot open the file (issue #78's three-removal repro on
    /// <c>StandardSalesInvoiceVatSpec.docx</c>'s first table). Neither keeping a cell there nor dropping
    /// the row is what the caller asked for, so the call is refused naming the row.
    /// <see cref="TableGridConsistencyGuard"/>'s <see cref="TableGridViolationKind.NoCells"/> check is the
    /// backstop behind this one.</item>
    /// </list>
    /// </remarks>
    private static void RejectUnremovableColumn(Table table, int tableIndex, int gridColumn, int gridCount)
    {
        var rows = TableGridNavigator.Rows(table);
        var rowsWithInnerRow = 0;
        var wellFormedRows = 0;
        int? firstMalformedRow = null;
        for (var r = 0; r < rows.Count; r++)
        {
            if (rows[r].InnerRow is not { } innerRow)
            {
                continue;
            }

            rowsWithInnerRow++;
            var coverage = TableGridNavigator.Coverage(innerRow);
            if (coverage.Total != gridCount)
            {
                firstMalformedRow ??= r;
                continue;
            }

            wellFormedRows++;
            var cells = TableGridNavigator.Cells(innerRow);
            if (cells.Count == 1 && cells[0].GridSpan == 1
                && gridColumn >= coverage.Before && gridColumn < coverage.ContentEnd)
            {
                throw new ArgumentException(
                    $"Cannot remove grid column {gridColumn} from table {tableIndex}: in row {r} it is that "
                    + $"row's ONLY cell (the rest of the row is skipped grid columns - w:gridBefore "
                    + $"{coverage.Before}, w:gridAfter {coverage.After}), and a table row with no cells makes "
                    + "the document impossible for Word to open. Pick a different column, or remove that row by "
                    + "hand first if it is no longer wanted.",
                    nameof(gridColumn));
            }
        }

        if (firstMalformedRow is { } malformed && wellFormedRows == 0)
        {
            throw new ArgumentException(
                $"Cannot remove a column from table {tableIndex}: none of its {rowsWithInnerRow} row(s) matches "
                + $"its w:tblGrid ({gridCount} declared column(s); row {malformed} covers "
                + $"{TableGridNavigator.Coverage(rows[malformed].InnerRow!).Total}), so there is no row the "
                + "removal could update and dropping a w:gridCol would only widen the mismatch. Repair the "
                + "table's w:tblGrid by hand before removing columns.",
                nameof(gridColumn));
        }

        if (rowsWithInnerRow == 0)
        {
            throw new ArgumentException(
                $"Cannot remove a column from table {tableIndex}: it has no row with cells to update, so "
                + "removing a w:gridCol would leave the grid describing columns no row has.",
                nameof(gridColumn));
        }
    }

    /// <summary>
    /// Scales every width except <paramref name="removedIndex"/> up so their sum equals the ORIGINAL
    /// total (largest-remainder rounding — the result sums exactly, no drift).
    /// </summary>
    private static int[] RedistributeWidth(int[] widthsBefore, int removedIndex)
    {
        var target = widthsBefore.Sum();
        var remaining = widthsBefore.Where((_, i) => i != removedIndex).ToArray();
        var remainingTotal = remaining.Sum();

        var exact = remaining.Select(w => (double)w * target / remainingTotal).ToArray();
        var result = exact.Select(e => (int)e).ToArray();
        var shortfall = target - result.Sum();
        foreach (var i in Enumerable.Range(0, exact.Length)
                     .OrderByDescending(i => exact[i] - result[i])
                     .Take(shortfall))
        {
            result[i]++;
        }

        return result;
    }

    /// <summary>
    /// Merges the contiguous physical cells <paramref name="fromColumn"/>..<paramref name="toColumn"/>
    /// (inclusive, 0-based within <paramref name="row"/>, the same cell indices <c>get_layout_info</c>
    /// reports) of one row into a single cell spanning all their grid columns (<c>w:gridSpan</c>). The first
    /// cell is kept (its content and binding survive) and widened; the rest are deleted. Horizontal merge
    /// only — the grid is unchanged, so the row's total grid coverage is preserved.
    /// </summary>
    /// <exception cref="ArgumentException">Range is invalid; an absorbed (non-first) cell holds a content control (would silently drop a binding); or the table uses vMerge.</exception>
    /// <exception cref="NotFoundException"><paramref name="tableIndex"/>/<paramref name="row"/> out of range.</exception>
    public static TableEditResult MergeCells(
        WordprocessingDocument doc, LayoutPart part, string? partName, int tableIndex, int row, int fromColumn, int toColumn)
    {
        ArgumentNullException.ThrowIfNull(doc);
        var (table, partFile, gridCount) = ResolveTable(doc, part, partName, tableIndex);
        RejectUnsupportedShape(table, "merge_cells");

        var (mergeRow, cells) = ResolveRowCells(table, tableIndex, partFile, row);
        var mergeCoverage = TableGridNavigator.Coverage(mergeRow);
        if (fromColumn < 0 || toColumn >= cells.Count || fromColumn >= toColumn)
        {
            throw new ArgumentException(
                $"merge_cells needs 0 <= fromColumn < toColumn < cell count ({cells.Count}) for table "
                + $"{tableIndex} row {row}; got fromColumn {fromColumn}, toColumn {toColumn}.",
                nameof(fromColumn));
        }

        for (var i = fromColumn + 1; i <= toColumn; i++)
        {
            if (cells[i].CellChild is SdtCell || cells[i].CellChild.Descendants<SdtElement>().Any())
            {
                throw new ArgumentException(
                    $"merge_cells would discard the content of cell {i} (table {tableIndex} row {row}), which "
                    + "holds a bound field/label control - merging drops the absorbed cells' content, which "
                    + "would silently lose that binding. Remove the control first (remove_control) if you "
                    + "really want to merge over it.",
                    nameof(toColumn));
            }
        }

        var newSpan = 0;
        for (var i = fromColumn; i <= toColumn; i++)
        {
            newSpan += cells[i].GridSpan;
        }

        SetCellGridSpan(cells[fromColumn].InnerCell, newSpan);
        var startGrid = mergeCoverage.Before + cells.Take(fromColumn).Sum(c => c.GridSpan);
        if (cells[fromColumn].InnerCell is not null)
        {
            SetCellWidth(cells[fromColumn].InnerCell!, SumGridWidths(table, startGrid, newSpan));
        }

        for (var i = toColumn; i > fromColumn; i--)
        {
            cells[i].CellChild.Remove();
        }

        return new TableEditResult
        {
            Operation = "merge_cells",
            Part = partFile,
            TableIndex = tableIndex,
            ColumnIndex = startGrid,
            RowsAffected = 1,
            ColumnCountBefore = gridCount,
            ColumnCountAfter = gridCount,
            Summary = $"Merged cells {fromColumn}..{toColumn} of table {tableIndex} row {row}"
                + (partFile == "document.xml" ? string.Empty : $" in {partFile}")
                + $" into one cell spanning {newSpan} grid column(s).",
        };
    }

    /// <summary>
    /// Splits a single spanned cell (<paramref name="column"/>, 0-based within <paramref name="row"/>) back
    /// into one cell per grid column it spans: the cell's <c>w:gridSpan</c> is reset to 1 (keeping its
    /// content) and empty unit cells are inserted after it for the remaining grid columns. Horizontal split
    /// only — the grid is unchanged, so the row's total grid coverage is preserved.
    /// </summary>
    /// <exception cref="ArgumentException">The cell is not spanned (nothing to split); or the table uses vMerge.</exception>
    /// <exception cref="NotFoundException"><paramref name="tableIndex"/>/<paramref name="row"/> out of range.</exception>
    public static TableEditResult SplitCell(
        WordprocessingDocument doc, LayoutPart part, string? partName, int tableIndex, int row, int cellIndex)
    {
        ArgumentNullException.ThrowIfNull(doc);
        var (table, partFile, gridCount) = ResolveTable(doc, part, partName, tableIndex);
        RejectUnsupportedShape(table, "split_cells");

        var (splitRow, cells) = ResolveRowCells(table, tableIndex, partFile, row);
        var splitCoverage = TableGridNavigator.Coverage(splitRow);
        if (cellIndex < 0 || cellIndex >= cells.Count)
        {
            throw new ArgumentException(
                $"cellIndex {cellIndex} is out of range for table {tableIndex} row {row}; it has {cells.Count} cell(s).",
                nameof(cellIndex));
        }

        var span = cells[cellIndex].GridSpan;
        if (span <= 1)
        {
            throw new ArgumentException(
                $"Cell {cellIndex} of table {tableIndex} row {row} is not spanned (gridSpan 1); there is nothing "
                + "to split.",
                nameof(cellIndex));
        }

        var startGrid = splitCoverage.Before + cells.Take(cellIndex).Sum(c => c.GridSpan);
        SetCellGridSpan(cells[cellIndex].InnerCell, 1);
        if (cells[cellIndex].InnerCell is not null)
        {
            SetCellWidth(cells[cellIndex].InnerCell!, GridWidthAt(table, startGrid));
        }

        OpenXmlElement anchor = cells[cellIndex].CellChild;
        for (var g = 1; g < span; g++)
        {
            var newCell = BuildCellLike(cells[cellIndex].InnerCell, new Paragraph(), GridWidthAt(table, startGrid + g));
            anchor.InsertAfterSelf(newCell);
            anchor = newCell;
        }

        return new TableEditResult
        {
            Operation = "split_cells",
            Part = partFile,
            TableIndex = tableIndex,
            ColumnIndex = startGrid,
            RowsAffected = 1,
            ColumnCountBefore = gridCount,
            ColumnCountAfter = gridCount,
            Summary = $"Split cell {cellIndex} of table {tableIndex} row {row}"
                + (partFile == "document.xml" ? string.Empty : $" in {partFile}")
                + $" into {span} single-column cell(s).",
        };
    }

    // ---- shared resolution / preconditions ----

    private static (TableRow InnerRow, IReadOnlyList<TableGridNavigator.CellSlot> Cells) ResolveRowCells(
        Table table, int tableIndex, string partFile, int row)
    {
        var rows = TableGridNavigator.Rows(table);
        if (row < 0 || row >= rows.Count)
        {
            throw new NotFoundException(
                $"Row index {row} is out of range for table {tableIndex}; it has {rows.Count} row(s).",
                NotFoundTarget.TableCoordinate);
        }

        var innerRow = rows[row].InnerRow
            ?? throw new NotFoundException(
                $"Table {tableIndex} row {row} resolves to a row-level control with no inner table row to operate on.",
                NotFoundTarget.TableCoordinate);
        return (innerRow, TableGridNavigator.Cells(innerRow));
    }

    /// <summary>
    /// Writes <paramref name="widths"/> (one per grid column) onto the table's <c>w:gridCol</c>s and
    /// re-slices every well-formed row's cell widths (<c>w:tcW</c>) to match — a spanning cell gets the
    /// sum of its covered grid columns. Rows whose grid coverage doesn't equal the grid (pre-existing
    /// malformed rows) are left alone rather than mis-sliced. Shared by <see cref="SetColumnWidths"/>
    /// and <see cref="RemoveColumn"/>'s width redistribution.
    /// </summary>
    private static int ApplyGridWidths(Table table, IReadOnlyList<int> widths)
    {
        var gridColumns = GridColumns(table);
        for (var i = 0; i < gridColumns.Count && i < widths.Count; i++)
        {
            gridColumns[i].Width = widths[i].ToString(CultureInfo.InvariantCulture);
        }

        var rowsAffected = 0;
        foreach (var row in TableGridNavigator.Rows(table))
        {
            if (row.InnerRow is not null && ResliceRowWidths(row.InnerRow, widths))
            {
                rowsAffected++;
            }
        }

        return rowsAffected;
    }

    /// <summary>
    /// Rewrites one row's cell widths (<c>w:tcW</c>) to the sum of the grid columns each cell covers.
    /// Returns false (touching nothing) when the row's coverage doesn't equal the grid — a pre-existing
    /// malformed row is left alone rather than mis-sliced.
    /// </summary>
    private static bool ResliceRowWidths(TableRow innerRow, IReadOnlyList<int> widths)
    {
        var cells = TableGridNavigator.Cells(innerRow);
        var coverage = TableGridNavigator.Coverage(innerRow);
        if (coverage.Total != widths.Count)
        {
            return false;
        }

        // The row's first physical cell starts at grid column `Before`, not 0 — the widths of any leading
        // skipped columns belong to no cell here.
        var cursor = coverage.Before;
        foreach (var cell in cells)
        {
            if (cell.InnerCell is not null)
            {
                var widthTwips = 0;
                for (var g = cursor; g < cursor + cell.GridSpan && g < widths.Count; g++)
                {
                    widthTwips += widths[g];
                }

                SetCellWidth(cell.InnerCell, widthTwips);
            }

            cursor += cell.GridSpan;
        }

        return true;
    }

    /// <summary>
    /// Reads every <c>w:gridCol</c> width; false when any is missing/unparsable/non-positive — then there
    /// is no reliable basis for proportional or per-column width arithmetic.
    /// </summary>
    private static bool TryGetGridWidths(Table table, out int[] widths)
    {
        var gridColumns = GridColumns(table);
        widths = new int[gridColumns.Count];
        for (var i = 0; i < gridColumns.Count; i++)
        {
            if (!int.TryParse(gridColumns[i].Width?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out widths[i])
                || widths[i] <= 0)
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>
    /// True when the slot is a PLAIN table cell with no visible content — no text, no content control,
    /// no drawing. A cell-level control wrapper (<c>SdtCell</c>) is never "empty": it carries a binding.
    /// </summary>
    private static bool IsEmptyCell(TableGridNavigator.CellSlot slot) =>
        slot.CellChild is TableCell cell
        && string.IsNullOrWhiteSpace(cell.InnerText)
        && !cell.Descendants<SdtElement>().Any()
        && !cell.Descendants<Drawing>().Any();

    private static int SumGridWidths(Table table, int startGrid, int span)
    {
        var widths = GridColumns(table);
        var sum = 0;
        for (var g = startGrid; g < startGrid + span && g < widths.Count; g++)
        {
            sum += int.TryParse(widths[g].Width?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var w) ? w : 0;
        }

        return sum == 0 ? DefaultColumnWidthTwips : sum;
    }

    private static int GridWidthAt(Table table, int gridIndex)
    {
        var widths = GridColumns(table);
        if (gridIndex >= 0 && gridIndex < widths.Count
            && int.TryParse(widths[gridIndex].Width?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var w) && w > 0)
        {
            return w;
        }

        return DefaultColumnWidthTwips;
    }



    private static (Table Table, string PartName, int GridCount) ResolveTable(
        WordprocessingDocument doc, LayoutPart part, string? partName, int tableIndex)
    {
        var (root, partFile) = LocationResolver.ResolvePart(doc, part, partName);
        var table = TableGridNavigator.TableAt(root, tableIndex, LocationResolver.DescribePart(partFile));
        var gridCount = TableGridNavigator.GridColumnCount(table);
        if (gridCount == 0)
        {
            throw new ArgumentException(
                $"Table {tableIndex} in {LocationResolver.DescribePart(partFile)} has no w:tblGrid (declared "
                + "columns); column operations need a grid to address columns against.",
                nameof(tableIndex));
        }

        return (table, partFile, gridCount);
    }

    /// <summary>
    /// Rejects the one remaining table shape whose column arithmetic is deferred pending real capture
    /// (GitHub issue #9): vertical merges (<c>w:vMerge</c>). Still absent from every layout reviewed, so this never
    /// blocks a real BC table; it stops the tools from guessing on a shape nobody has verified.
    /// </summary>
    /// <remarks>
    /// Leading/trailing skipped grid cells (<c>w:gridBefore</c>/<c>w:gridAfter</c>) used to be rejected here
    /// too, on the same "absent from the corpus" grounds. They are not absent: the 2026-08-01 layout review
    /// found <c>gridAfter</c> in seven base-app layouts, including on the LINE-ITEMS repeater table of
    /// <c>StandardSalesInvoiceVatSpec.docx</c> — so the rejection blocked editing the lines table of a stock
    /// BC sales invoice. Every operation now carries the offset instead; see
    /// <see cref="TableGridNavigator.RowGridCoverage"/>. <c>gridBefore</c> is still unseen in the wild but
    /// falls out of the same arithmetic, so it is supported rather than specially rejected.
    /// </remarks>
    private static void RejectUnsupportedShape(Table table, string op)
    {
        foreach (var row in TableGridNavigator.Rows(table))
        {
            if (row.InnerRow is null)
            {
                continue;
            }

            foreach (var cell in TableGridNavigator.Cells(row.InnerRow))
            {
                if (cell.InnerCell?.GetFirstChild<TableCellProperties>()?.GetFirstChild<VerticalMerge>() is not null)
                {
                    throw new ArgumentException(
                        $"{op} does not support tables that use vertical cell merges (w:vMerge); this shape is "
                        + "absent from every layout reviewed, so there is no fixture to implement the span "
                        + "arithmetic against; author a reference vertical-merge table in Word first "
                        + "(GitHub issue #9).");
                }
            }
        }
    }

    // ---- cell / grid building blocks ----

    private static bool HasTblHeader(TableRow? innerRow) =>
        innerRow?.GetFirstChild<TableRowProperties>()?.GetFirstChild<TableHeader>() is not null;

    private static Paragraph BuildHeaderParagraph(DatasetTree? schema, InsertColumnOptions options, Func<int> nextId)
    {
        if (schema is not null && !string.IsNullOrWhiteSpace(options.HeaderLabelPath))
        {
            return new Paragraph(SdtFactory.BuildLabel(schema, options.HeaderLabelPath!, id: nextId()));
        }

        var text = !string.IsNullOrWhiteSpace(options.HeaderText)
            ? options.HeaderText!
            : options.DataPath is { } dp
                ? SdtFactory.Humanize(dp.Split('/', StringSplitOptions.RemoveEmptyEntries)[^1])
                : string.Empty;

        return text.Length == 0
            ? new Paragraph()
            : new Paragraph(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
    }

    /// <summary>
    /// Builds a new single-column cell that visually belongs next to <paramref name="neighbor"/>: the
    /// neighbor's <c>w:tcPr</c> (borders, shading, vertical alignment, margins — minus its span/merge
    /// state), its first paragraph's <c>w:pPr</c> (justification, spacing, style), and its first styled
    /// run's <c>w:rPr</c> (applied to <paramref name="paragraph"/>'s runs and any sdt's own run
    /// properties) are cloned onto the new cell. With no neighbor (or an unformatted one) the cell is
    /// plain — exactly the pre-clone behavior.
    /// </summary>
    private static TableCell BuildCellLike(TableCell? neighbor, Paragraph paragraph, int width)
    {
        var cell = new TableCell();
        if (neighbor?.GetFirstChild<TableCellProperties>()?.CloneNode(true) is TableCellProperties tcPr)
        {
            // Span/merge state describes the NEIGHBOR's grid position, never the new unit cell's.
            tcPr.RemoveAllChildren<GridSpan>();
            tcPr.RemoveAllChildren<HorizontalMerge>();
            tcPr.RemoveAllChildren<VerticalMerge>();
            cell.AppendChild(tcPr);
        }

        if (paragraph.ParagraphProperties is null
            && neighbor?.Descendants<Paragraph>().FirstOrDefault()?.ParagraphProperties is { } pPr)
        {
            paragraph.ParagraphProperties = (ParagraphProperties)pPr.CloneNode(true);
        }

        if (neighbor?.Descendants<Run>().FirstOrDefault(r => r.RunProperties is not null)?.RunProperties is { } rPr)
        {
            foreach (var run in paragraph.Descendants<Run>().Where(r => r.RunProperties is null))
            {
                run.RunProperties = (RunProperties)rPr.CloneNode(true);
            }

            // An sdt's own w:sdtPr/w:rPr is what Word styles a control's placeholder/content with; the
            // corpus-verified child order puts rPr first (see SdtFactory.BuildRepeaterTable's remarks).
            foreach (var sdtPr in paragraph.Descendants<SdtProperties>()
                         .Where(p => p.GetFirstChild<RunProperties>() is null))
            {
                sdtPr.PrependChild((RunProperties)rPr.CloneNode(true));
            }
        }

        cell.AppendChild(paragraph);
        SetCellWidth(cell, width);
        return cell;
    }

    private static void SetCellWidth(TableCell cell, int widthTwips)
    {
        var tcPr = EnsureTcPr(cell);
        var tcW = tcPr.GetFirstChild<TableCellWidth>();
        if (tcW is null)
        {
            tcW = new TableCellWidth();
            InsertTcPrChildInOrder(tcPr, tcW);
        }

        tcW.Type = TableWidthUnitValues.Dxa;
        tcW.Width = widthTwips.ToString(CultureInfo.InvariantCulture);
    }

    private static void SetCellGridSpan(TableCell? cell, int span)
    {
        if (cell is null)
        {
            return;
        }

        var tcPr = EnsureTcPr(cell);
        var gridSpan = tcPr.GetFirstChild<GridSpan>();
        if (span <= 1)
        {
            gridSpan?.Remove(); // span of 1 is the default — no element needed.
            return;
        }

        if (gridSpan is null)
        {
            gridSpan = new GridSpan();
            InsertTcPrChildInOrder(tcPr, gridSpan);
        }

        gridSpan.Val = span;
    }

    private static TableCellProperties EnsureTcPr(TableCell cell)
    {
        var tcPr = cell.GetFirstChild<TableCellProperties>();
        if (tcPr is null)
        {
            tcPr = new TableCellProperties();
            cell.PrependChild(tcPr);
        }

        return tcPr;
    }

    /// <summary>
    /// Inserts <paramref name="child"/> into <paramref name="tcPr"/> after the LAST present element that
    /// precedes it in the CT_TcPr schema sequence (<c>w:cnfStyle</c>, <c>w:tcW</c>, <c>w:gridSpan</c>, …),
    /// else as the first child. Keeps the property order valid so the pre-save OpenXmlValidator gate stays
    /// clean. The predecessor set depends on WHICH child is being inserted: a <c>w:gridSpan</c> must land
    /// after an existing <c>w:tcW</c> — inserting it at the front of a tcPr that already carries a width
    /// (the shape of every real BC table cell) is a schema violation the validator gate then rejects,
    /// which is exactly how merge_cells failed on the corpus quote's spacer row (2026-07-31 scenario e2e).
    /// <see cref="TcPrSequence"/> is the schema order itself, so this generalizes to any child rather than
    /// special-casing the two the editor happened to insert first (a <c>w:tcBorders</c> must land after
    /// <c>w:tcW</c>/<c>w:gridSpan</c>/the merge elements, for instance).
    /// </summary>
    private static void InsertTcPrChildInOrder(TableCellProperties tcPr, OpenXmlElement child)
    {
        var childRank = Array.IndexOf(TcPrSequence, child.GetType());

        OpenXmlElement? predecessor = null;
        if (childRank >= 0)
        {
            foreach (var existing in tcPr.ChildElements)
            {
                var rank = Array.IndexOf(TcPrSequence, existing.GetType());
                if (rank >= 0 && rank < childRank)
                {
                    predecessor = existing;
                }
            }
        }

        if (predecessor is not null)
        {
            tcPr.InsertAfter(child, predecessor);
        }
        else
        {
            tcPr.InsertAt(child, 0);
        }
    }

    /// <summary>
    /// The CT_TcPr child sequence, in schema order — the authority <see cref="InsertTcPrChildInOrder"/>
    /// places a new property against. Only the elements this editor can encounter or insert are listed;
    /// anything absent from it is treated as "no known position" and inserted first, exactly as before.
    /// </summary>
    private static readonly Type[] TcPrSequence =
    {
        typeof(ConditionalFormatStyle),
        typeof(TableCellWidth),
        typeof(GridSpan),
        typeof(HorizontalMerge),
        typeof(VerticalMerge),
        typeof(TableCellBorders),
        typeof(Shading),
        typeof(NoWrap),
        typeof(TableCellMargin),
        typeof(TextDirection),
        typeof(TableCellFitText),
        typeof(TableCellVerticalAlignment),
        typeof(HideMark),
    };

    // "Unexpected" is the operative word: ResolveTable already rejects (with an ArgumentException) any table
    // whose grid column count is 0 before any caller reaches here, so a missing w:tblGrid at this point is an
    // internal invariant violation, not a lookup failure a caller could fix by retrying - left as a plain
    // InvalidOperationException (→ internal_error), not NotFoundException.
    private static TableGrid GridElement(Table table) =>
        table.GetFirstChild<TableGrid>()
            ?? throw new InvalidOperationException("Table has no w:tblGrid (unexpected — checked by ResolveTable).");

    private static List<GridColumn> GridColumns(Table table) => GridElement(table).Elements<GridColumn>().ToList();

    private static int MeanGridWidth(Table table)
    {
        var widths = GridColumns(table)
            .Select(g => int.TryParse(g.Width?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var w) ? w : (int?)null)
            .Where(w => w.HasValue)
            .Select(w => w!.Value)
            .ToList();
        return widths.Count == 0 ? DefaultColumnWidthTwips : (int)Math.Round(widths.Average());
    }

    private static void SyncTableWidthToGrid(Table table)
    {
        var tblW = table.GetFirstChild<TableProperties>()?.TableWidth;
        if (tblW?.Type is null || tblW.Type.Value != TableWidthUnitValues.Dxa)
        {
            return; // Only a fixed (dxa) total width can be kept in sync; auto/pct are recomputed by Word.
        }

        var sum = GridColumns(table).Sum(g => int.TryParse(g.Width?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var w) ? w : 0);
        tblW.Width = sum.ToString(CultureInfo.InvariantCulture);
    }

    private static string DescribeMode(InsertColumnOptions options) => options.Mode switch
    {
        InsertColumnMode.Field => $"bound-field ('{options.DataPath}')",
        InsertColumnMode.Label => $"bound-label ('{options.DataPath}')",
        _ => "plain-text",
    };
}
