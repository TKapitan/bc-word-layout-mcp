using System.ComponentModel;
using System.Globalization;
using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;
using ModelContextProtocol.Server;
using static BcWordLayout.McpHost.Tools.ToolGuards;

namespace BcWordLayout.McpHost.Tools;

/// <summary>
/// MCP tools that add or change a table's STRUCTURE (as opposed to <see cref="EditTools"/>, which edits an
/// existing control's or plain-text cell's content): authoring a whole new repeater table
/// (<c>insert_repeater_table</c>), resizing columns (<c>set_column_widths</c>), adding/removing a column
/// (<c>insert_column</c>/<c>remove_column</c>), and horizontally merging/splitting cells within a row
/// (<c>merge_cells</c>/<c>split_cells</c>). <c>insert_repeater_table</c> sits here rather than in
/// <see cref="EditTools"/> because it is fundamentally a table-shape operation (it builds the table's grid
/// and rows from scratch), even though its edit-safety wrapper is the content-control one
/// (<see cref="ToolGuards.GuardEdit"/>) — the other five route through
/// <see cref="ToolGuards.GuardTableEdit"/>. Both wrappers, and the grid-consistency backstop
/// <see cref="ToolGuards.GuardMutate{TResult}"/> runs for every mutating tool, live in <see cref="ToolGuards"/>,
/// which also holds the single shared per-path edit lock these tools serialize against alongside every
/// other tool family's mutating tools.
/// </summary>
[McpServerToolType]
public static class TableTools
{
    [McpServerTool(Name = "insert_repeater_table")]
    [Description("Insert a complete repeater TABLE bound to a repeating data item (e.g. '/Header/Line', as "
                 + "reported by list_dataset_fields): a header row (label controls where a label column "
                 + "exists for a given column - by default suffixed Lbl/_Lbl, see list_dataset_fields's "
                 + "isLabel flag - else static humanized text) plus one data row wrapped in a "
                 + "repeating-section control, with one bound field control per column. This is the flagship "
                 + "tool for adding a new line-items-style table to a layout. By default the table gets the "
                 + "BC-NATIVE look (look='bc'): no drawn grid, just a rule under the header row, exactly like "
                 + "every real BC lines table - add the rule above a totals block with set_cell_borders. "
                 + "The result reports tableIndex and dataRowIndex - the new table's own 0-based index and "
                 + "the 0-based index of its repeating DATA row - so follow-up edits need no re-read. "
                 + "NESTING: for per-line detail (components, serial/lot nos), the STANDARD BC shape is a "
                 + "detail ROW under the line row - use insert_repeater_row with this call's returned "
                 + "controlId. Hosting a whole nested TABLE inside a cell (locationType='tableCell' with "
                 + "this call's tableIndex/dataRowIndex) also works but is NOT how standard BC documents "
                 + "are laid out - prefer insert_repeater_row. "
                 + "v1 SCOPE: only supports the main "
                 + "document body (layoutPart='body', the only supported value) - a repeater table in a "
                 + "header/footer is explicitly deferred (unlike insert_field/insert_label, which do support "
                 + "layoutPart='header'/'footer'). Same write/validate/save-or-reject safety and response "
                 + "shape as insert_field.")]
    public static ToolResponse InsertRepeaterTable(
        [Description("Absolute path to the .docx layout file.")] string layoutPath,
        [Description("Dataset path to a repeating data item, e.g. '/Header/Line' (see list_dataset_fields).")]
        string dataItem,
        [Description("Comma-separated leaf column names of that data item, in order, e.g. "
                     + "'ItemNo_Line,Description_Line,Quantity_Line'.")] string columns,
        [Description("Where to insert: 'documentEnd', 'afterControl', 'tableCell', or 'atText' "
                     + "(case-insensitive).")] string locationType,
        [Description("Required for 'afterControl': the w:id of the control to insert after.")] int? controlId = null,
        [Description("Required for 'tableCell': 0-based table index in document order.")] int? tableIndex = null,
        [Description("Required for 'tableCell': 0-based row index within the table.")] int? row = null,
        [Description("Required for 'tableCell': 0-based cell/column index within the row.")] int? col = null,
        [Description("Required for 'atText': substring to search for in existing run text (ordinal match).")]
        string? searchText = null,
        [Description("When true (default), each header cell binds to its column's label column when one can "
                     + "be found; otherwise every header cell is static humanized column-name text.")]
        bool headerFromLabels = true,
        [Description("Optional Word table style name to reference via w:tblStyle, e.g. 'TableGrid'.")]
        string? tableStyle = null,
        [Description("Optional comma-separated column widths in twips (1/20 pt), one per column; count must "
                     + "match the number of columns. Omit for an even default width per column.")]
        string? columnWidths = null,
        [Description("Optional comma-separated per-column alignments ('left'/'center'/'right', one per "
                     + "column), applied to both the header and data cell of each column - real BC line "
                     + "tables right-align their quantity/price/amount columns. Omit for the default (left).")]
        string? columnAlignments = null,
        [Description("Border treatment: 'bc' (default) draws NO table grid, only a 1/2-pt rule under the "
                     + "header row - the corpus-verified look of every real BC lines table; 'grid' draws an "
                     + "explicit single-line border on every edge and between every cell. Use "
                     + "set_cell_borders afterwards for the other rules a BC document has (e.g. above a "
                     + "totals block).")] string look = "bc",
        [Description("v1 SCOPE: only 'body' (the default) is supported for a repeater table - passing "
                     + "'header'/'footer' here is rejected with invalid_argument (repeaters in headers/"
                     + "footers are deferred - GitHub issue #10). Present for parameter-shape consistency "
                     + "with insert_field/insert_label, which DO support 'header'/'footer'.")]
        string layoutPart = "body",
        [Description("Not supported for insert_repeater_table (see layoutPart) - always omit this.")]
        string? partName = null)
    {
        return GuardEdit(layoutPath, doc =>
        {
            var columnList = SplitCommaList(columns);
            if (columnList.Length == 0)
            {
                throw new ArgumentException("At least one column is required (columns was empty).", nameof(columns));
            }

            var options = new RepeaterTableOptions
            {
                Look = ParseTableBorderLookOrThrow(look),
                HeaderFromLabels = headerFromLabels,
                TableStyle = string.IsNullOrWhiteSpace(tableStyle) ? null : tableStyle,
                ColumnWidths = ParseColumnWidths(columnWidths),
                ColumnAlignments = string.IsNullOrWhiteSpace(columnAlignments) ? null : SplitCommaList(columnAlignments),
            };

            var location = BuildLocation(locationType, controlId, tableIndex, row, col, searchText, layoutPart, partName);
            return LayoutEditor.InsertRepeaterTable(doc, dataItem, columnList, location, options);
        });
    }

    [McpServerTool(Name = "insert_repeater_row")]
    [Description("Add a nested DETAIL ROW repeater inside an EXISTING repeater's item - the STANDARD BC "
                 + "shape for per-line detail: 'another line with the required details' rendered UNDER each "
                 + "line row (assembly components, serial nos, lot nos), repeating once per parent row. "
                 + "This is how real BC documents nest - NOT a table hosted in a side column. "
                 + "parentControlId is the outer repeater's controlId (returned by insert_repeater_table; "
                 + "also in get_layout_info's control inventory). dataItem must be a repeating data item "
                 + "DIRECTLY under the parent's (e.g. '/Header/Line/AssemblyLine' under '/Header/Line'); for "
                 + "deeper levels, call this again on the child repeater's own controlId. cells lays the row "
                 + "out on the PARENT table's grid, comma-separated, one entry per cell: '-' is an empty "
                 + "spacer cell, a name is a leaf column of the child item (label-shaped names become label "
                 + "controls), 'a+b' chains several columns inline in ONE cell (the corpus shape for "
                 + "'label: value' pairs - note they render with no separator between them), and an optional "
                 + "'N:' prefix makes the cell span N grid columns (e.g. '-,3:Description_AssemblyLine,"
                 + "2:Quantity_AssemblyLine,2:-'). The spans MUST sum to the parent table's columnCount. "
                 + "The row is appended after the line row and any existing detail rows, so repeated calls "
                 + "stack detail lines in order. Same write/validate/save-or-reject safety as insert_field, "
                 + "plus the grid-consistency backstop.")]
    public static ToolResponse InsertRepeaterRow(
        [Description("Absolute path to the .docx layout file.")] string layoutPath,
        [Description("The OUTER repeater's controlId (from insert_repeater_table's result or "
                     + "get_layout_info).")] int parentControlId,
        [Description("Dataset path of the repeating data item DIRECTLY under the parent's, e.g. "
                     + "'/Header/Line/AssemblyLine' (see list_dataset_fields).")] string dataItem,
        [Description("Comma-separated cell specs laying the detail row on the parent grid: '-' = empty "
                     + "spacer, 'Name' = one bound column, 'a+b' = several columns inline in one cell, "
                     + "optional 'N:' span prefix. Spans must sum to the parent table's columnCount.")]
        string cells,
        [Description("Optional comma-separated per-cell alignments ('left'/'center'/'right', or '-' to keep "
                     + "the default), one per cell spec.")] string? alignments = null)
    {
        return GuardEdit(layoutPath, doc =>
            LayoutEditor.InsertRepeaterRow(doc, parentControlId, dataItem, ParseRepeaterRowCells(cells, alignments)));
    }

    /// <summary>
    /// Parses <c>insert_repeater_row</c>'s cells DSL (<c>[N:]name(+name)*</c> or <c>-</c>, comma-separated)
    /// plus the optional per-cell alignment list into <see cref="RepeaterRowCell"/>s. Malformed input throws
    /// <see cref="ArgumentException"/> (mapped to <c>invalid_argument</c>); semantic validation (columns
    /// exist, spans cover the grid) is the domain layer's.
    /// </summary>
    private static List<RepeaterRowCell> ParseRepeaterRowCells(string cells, string? alignments)
    {
        var specs = SplitCommaList(cells);
        if (specs.Length == 0)
        {
            throw new ArgumentException("cells must contain at least one cell spec.", nameof(cells));
        }

        var alignmentList = string.IsNullOrWhiteSpace(alignments) ? null : SplitCommaList(alignments);
        if (alignmentList is not null && alignmentList.Length != specs.Length)
        {
            throw new ArgumentException(
                $"alignments has {alignmentList.Length} entr{(alignmentList.Length == 1 ? "y" : "ies")} but "
                + $"cells has {specs.Length}; supply exactly one per cell (use '-' to keep a cell's default).",
                nameof(alignments));
        }

        var result = new List<RepeaterRowCell>(specs.Length);
        for (var i = 0; i < specs.Length; i++)
        {
            var spec = specs[i];
            var span = 1;
            var colonAt = spec.IndexOf(':');
            if (colonAt > 0)
            {
                if (!int.TryParse(spec[..colonAt], NumberStyles.Integer, CultureInfo.InvariantCulture, out span))
                {
                    throw new ArgumentException(
                        $"cells entry '{spec}' has a malformed span prefix; use e.g. '3:Description_AssemblyLine'.",
                        nameof(cells));
                }

                spec = spec[(colonAt + 1)..].Trim();
            }

            var alignment = alignmentList?[i] is { } a && a != "-" ? a : null;
            result.Add(new RepeaterRowCell
            {
                Span = span,
                Columns = spec == "-" || spec.Length == 0
                    ? []
                    : spec.Split('+', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries),
                Alignment = alignment,
            });
        }

        return result;
    }

    [McpServerTool(Name = "insert_table")]
    [Description("Insert a NEW plain (unbound) table - the building block for every NON-repeating layout "
                 + "section a real BC document is made of: side-by-side address columns, label/value "
                 + "header-info grids, right-anchored totals blocks. (For the LINE-ITEMS table bound to a "
                 + "repeating data item, use insert_repeater_table instead.) The table is rows x columns of "
                 + "empty single-column cells, borderless by default (the corpus shape for those blocks; pass "
                 + "withBorders=true for the same explicit single-line grid insert_repeater_table draws). "
                 + "columnWidths (twips, comma-separated, one per column) defaults to an even split of the "
                 + "full A4 content width (10206); columnAlignments (left/center/right, comma-separated, one "
                 + "per column) seeds each cell's paragraph justification - set_cell_text and "
                 + "insert_field/insert_label (locationType 'tableCell') preserve it afterwards. The result's "
                 + "tableIndex is the new table's 0-based document-order index: EXACTLY the index to pass to "
                 + "set_cell_text/insert_field/insert_label/set_column_widths next to fill the cells. v1 "
                 + "SCOPE: body only (like insert_repeater_table). Same write/validate/save-or-reject safety "
                 + "as insert_field.")]
    public static ToolResponse InsertTable(
        [Description("Absolute path to the .docx layout file.")] string layoutPath,
        [Description("Number of rows (1-100).")] int rows,
        [Description("Number of columns (1-30).")] int columns,
        [Description("Where to insert: 'documentEnd', 'afterControl', 'tableCell', or 'atText' "
                     + "(case-insensitive).")] string locationType,
        [Description("Required for 'afterControl': the w:id of the control to insert after.")] int? controlId = null,
        [Description("Required for 'tableCell': 0-based table index in document order.")] int? tableIndex = null,
        [Description("Required for 'tableCell': 0-based row index within the table.")] int? row = null,
        [Description("Required for 'tableCell': 0-based cell/column index within the row.")] int? col = null,
        [Description("Required for 'atText': substring to search for in existing run text (ordinal match).")]
        string? searchText = null,
        [Description("Optional comma-separated column widths in twips, one per column; omit for an even "
                     + "split of the full content width (10206 twips).")] string? columnWidths = null,
        [Description("Optional comma-separated per-column alignments ('left'/'center'/'right', one per "
                     + "column) seeding each cell's paragraph justification; omit for the default (left).")]
        string? columnAlignments = null,
        [Description("When true, draw the same explicit single-line borders insert_repeater_table draws; "
                     + "default false (borderless - the shape of BC's own address/info/totals blocks).")]
        bool withBorders = false,
        [Description("v1 SCOPE: only 'body' (the default) is supported - 'header'/'footer' are rejected "
                     + "with invalid_argument (same scope rule as insert_repeater_table).")]
        string layoutPart = "body",
        [Description("Not supported for insert_table (see layoutPart) - always omit this.")]
        string? partName = null)
    {
        return GuardTableEdit(layoutPath, doc =>
        {
            var location = BuildLocation(locationType, controlId, tableIndex, row, col, searchText, layoutPart, partName);
            var alignments = string.IsNullOrWhiteSpace(columnAlignments)
                ? null
                : SplitCommaList(columnAlignments);
            return LayoutEditor.InsertPlainTable(
                doc, rows, columns, location, ParseColumnWidths(columnWidths), alignments, withBorders);
        });
    }

    [McpServerTool(Name = "set_column_widths")]
    [Description("Set the column widths of a table, addressed by its 0-based table index (the SAME index "
                 + "get_layout_info reports in its tables[] section). widths is a comma-separated list of "
                 + "twips (1/20 pt), ONE PER GRID COLUMN - the count must equal the table's columnCount from "
                 + "get_layout_info. Each grid column's w:gridCol width is set, and every cell's width (w:tcW) "
                 + "is set to the sum of the grid-column widths it spans, so a cell that spans multiple columns "
                 + "(w:gridSpan - pervasive in real BC tables) stays as wide as the columns beneath it. If the "
                 + "table has a fixed total width (w:tblW type dxa) it is kept in sync with the new grid. "
                 + "REJECTED for tables that use vertical merges (w:vMerge) - that shape is not supported yet. "
                 + "Rows with leading/trailing skipped grid cells (w:gridBefore/w:gridAfter) ARE supported. "
                 + "To target a header/footer "
                 + "table, pass layoutPart='header'/'footer' (optionally partName). Same write/validate/"
                 + "save-or-reject safety as insert_field; additionally rejected if it would desync any row's "
                 + "cells from the grid.")]
    public static ToolResponse SetColumnWidths(
        [Description("Absolute path to the .docx layout file.")] string layoutPath,
        [Description("0-based table index in document order (see get_layout_info's tables[]).")] int tableIndex,
        [Description("Comma-separated column widths in twips, exactly one per GRID column (must match the "
                     + "table's columnCount).")] string widths,
        [Description("Which OOXML part the table is in: 'body' (default), 'header', or 'footer'.")] string layoutPart = "body",
        [Description("Only used when layoutPart is 'header'/'footer': a specific part file name (e.g. "
                     + "'header2.xml'); omit to target the DEFAULT header/footer part (see get_layout_info's "
                     + "partDetails for each part's role).")] string? partName = null)
    {
        return GuardTableEdit(layoutPath, doc =>
        {
            var part = ParseLayoutPartOrThrow(layoutPart);
            var widthList = ParseIntListOrThrow(widths, nameof(widths));
            return TableStructureEditor.SetColumnWidths(doc, part, partName, tableIndex, widthList);
        });
    }

    [McpServerTool(Name = "insert_column")]
    [Description("Append a NEW column to an existing table (line-items/repeater table or a plain top-section "
                 + "table), addressed by its 0-based table index (see get_layout_info's tables[]). Rows shaped "
                 + "like a right-anchored totals/summary block (leading EMPTY cells, content hugging the right "
                 + "edge - e.g. 'Total Incl VAT | 822.97' below the line rows) get NO new cell; their trailing "
                 + "content cell is widened across the new column instead, so the summary block keeps its exact "
                 + "look with its amounts still on the table's right edge. "
                 + "atColumn is the 0-based GRID position for the new column: omit it (or pass the table's "
                 + "columnCount) to append at the far-right edge, or give an INTERIOR position to insert "
                 + "between existing columns. An interior position must land on a cell BOUNDARY in every row "
                 + "that carries the new column's content (a header row, and the bound data row); a row with "
                 + "nothing there - a spacer, or a totals block's leading empty run - has its spanning cell "
                 + "widened instead of gaining a cell, so the block stays put. If it falls INSIDE a spanned "
                 + "cell of a content row the call is refused, naming the cell, rather than guessing (use "
                 + "split_cells there first). mode selects what the new column holds: 'field' or "
                 + "'label' bind the data cell via dataPath (a full dataset path exactly like insert_field/"
                 + "insert_label, e.g. '/Header/Line/Discount_Line' - it may bind a column of the repeater's "
                 + "data item OR a parent/Header field); 'plainText' adds an unbound column (headerText plus "
                 + "empty cells). The bound control is placed in the repeater's DATA row (or, in a table with "
                 + "no repeater, in every non-header row); header rows (w:tblHeader) get a header cell (bind "
                 + "headerLabelPath, else headerText, else a humanized column name); all other rows get an "
                 + "empty cell. One w:gridCol is added; width defaults to the mean of the existing columns. "
                 + "REJECTED for tables using w:vMerge (w:gridBefore/w:gridAfter rows are supported). Same write/validate/"
                 + "save-or-reject safety as insert_field, plus the grid-consistency backstop.")]
    public static ToolResponse InsertColumn(
        [Description("Absolute path to the .docx layout file.")] string layoutPath,
        [Description("0-based table index in document order (see get_layout_info's tables[]).")] int tableIndex,
        [Description("What the new column holds: 'field' (bound field), 'label' (bound label), or 'plainText' "
                     + "(unbound).")] string mode,
        [Description("For mode 'field'/'label': the full dataset path to bind the data cell to (e.g. "
                     + "'/Header/Line/Discount_Line'), exactly as insert_field/insert_label take. Ignored for "
                     + "'plainText'.")] string? dataPath = null,
        [Description("Optional header-cell text (used as the header in header rows; the only text for "
                     + "'plainText'). Omit to humanize the dataPath leaf name.")] string? headerText = null,
        [Description("Optional: bind the header cell to this label-column dataset path (by default a "
                     + "'*Lbl'/'*_Lbl' path - see list_dataset_fields's isLabel flag) instead of static text.")]
        string? headerLabelPath = null,
        [Description("Optional 0-based GRID position for the new column (0..columnCount). Omit, or pass the "
                     + "table's columnCount, to append at the far-right edge; an interior value inserts "
                     + "between existing columns (see this tool's description for the spanned-cell rule).")]
        int? atColumn = null,
        [Description("Optional new-column width in twips; omit for the mean of the existing columns.")]
        int? width = null,
        [Description("Which OOXML part the table is in: 'body' (default), 'header', or 'footer'.")] string layoutPart = "body",
        [Description("Only used when layoutPart is 'header'/'footer': a specific part file name (e.g. "
                     + "'header2.xml'); omit to target the DEFAULT header/footer part (see get_layout_info's "
                     + "partDetails for each part's role).")] string? partName = null)
    {
        return GuardTableEdit(layoutPath, doc =>
        {
            var part = ParseLayoutPartOrThrow(layoutPart);
            var options = new InsertColumnOptions
            {
                Mode = ParseInsertColumnModeOrThrow(mode),
                DataPath = string.IsNullOrWhiteSpace(dataPath) ? null : dataPath,
                HeaderText = string.IsNullOrWhiteSpace(headerText) ? null : headerText,
                HeaderLabelPath = string.IsNullOrWhiteSpace(headerLabelPath) ? null : headerLabelPath,
                Width = width,
            };
            return TableStructureEditor.InsertColumn(doc, part, partName, tableIndex, atColumn, options);
        });
    }

    [McpServerTool(Name = "remove_column")]
    [Description("Remove a whole column from a table, addressed by its 0-based table index and 0-based GRID "
                 + "column index (see get_layout_info's tables[]). The matching w:gridCol is removed, and in "
                 + "every row the cell covering that grid column is deleted when it spans only that column - "
                 + "including a bound field/label cell, whose binding is dropped with it (UNLIKE remove_control, "
                 + "which preserves the cell) - or has its w:gridSpan decremented when it spans more. Repeater "
                 + "and cell-level wrappers are preserved (only the inner cells/rows change). The last "
                 + "remaining column cannot be removed. The removed column's width is redistributed "
                 + "proportionally across the remaining columns, so a table that spanned the full content "
                 + "width still does (follow up with set_column_widths if you want a different distribution). "
                 + "REJECTED for tables using w:vMerge (w:gridBefore/"
                 + "w:gridAfter rows are supported). This is the tool for requests like 'remove the GST Amount column from the "
                 + "lines'. Same write/validate/save-or-reject safety as insert_field, plus the "
                 + "grid-consistency backstop that rejects any edit desyncing a row from the grid.")]
    public static ToolResponse RemoveColumn(
        [Description("Absolute path to the .docx layout file.")] string layoutPath,
        [Description("0-based table index in document order (see get_layout_info's tables[]).")] int tableIndex,
        [Description("0-based GRID column index to remove (0..columnCount-1 from get_layout_info).")] int column,
        [Description("Which OOXML part the table is in: 'body' (default), 'header', or 'footer'.")] string layoutPart = "body",
        [Description("Only used when layoutPart is 'header'/'footer': a specific part file name (e.g. "
                     + "'header2.xml'); omit to target the DEFAULT header/footer part (see get_layout_info's "
                     + "partDetails for each part's role).")] string? partName = null)
    {
        return GuardTableEdit(layoutPath, doc =>
        {
            var part = ParseLayoutPartOrThrow(layoutPart);
            return TableStructureEditor.RemoveColumn(doc, part, partName, tableIndex, column);
        });
    }

    [McpServerTool(Name = "merge_cells")]
    [Description("Horizontally merge a run of adjacent cells in ONE row into a single cell, addressed by the "
                 + "0-based table index, row index, and the 0-based physical cell indices fromColumn..toColumn "
                 + "(inclusive) within that row - the SAME table/row/cell indices get_layout_info reports. The "
                 + "first cell is KEPT (its content and any binding survive) and widened to span all the merged "
                 + "columns (w:gridSpan); the rest are deleted. The table grid is unchanged. REJECTED if any "
                 + "absorbed (non-first) cell holds a bound field/label control (merging would silently drop "
                 + "that binding - remove_control it first), or if the table uses w:vMerge. This is HORIZONTAL merge only; vertical merges (w:vMerge) are not supported "
                 + "yet. Same write/validate/save-or-reject safety as insert_field, plus the grid-consistency "
                 + "backstop.")]
    public static ToolResponse MergeCells(
        [Description("Absolute path to the .docx layout file.")] string layoutPath,
        [Description("0-based table index in document order (see get_layout_info's tables[]).")] int tableIndex,
        [Description("0-based row index within the table.")] int row,
        [Description("0-based physical cell index of the FIRST cell to merge (kept).")] int fromColumn,
        [Description("0-based physical cell index of the LAST cell to merge (inclusive); must be > fromColumn.")] int toColumn,
        [Description("Which OOXML part the table is in: 'body' (default), 'header', or 'footer'.")] string layoutPart = "body",
        [Description("Only used when layoutPart is 'header'/'footer': a specific part file name; omit for the "
                     + "first.")] string? partName = null)
    {
        return GuardTableEdit(layoutPath, doc =>
        {
            var part = ParseLayoutPartOrThrow(layoutPart);
            return TableStructureEditor.MergeCells(doc, part, partName, tableIndex, row, fromColumn, toColumn);
        });
    }

    [McpServerTool(Name = "split_cells")]
    [Description("Horizontally split ONE spanned cell back into one cell per grid column it spans, addressed "
                 + "by the 0-based table index, row index, and 0-based physical cell index within that row "
                 + "(the SAME indices get_layout_info reports). The cell's w:gridSpan is reset to 1 (keeping "
                 + "its content in the first resulting cell) and empty single-column cells are inserted after "
                 + "it for the remaining columns. The table grid is unchanged. REJECTED if the cell is not "
                 + "spanned (gridSpan 1 - nothing to split), or if the table uses w:vMerge. "
                 + "HORIZONTAL split only. Same write/validate/save-or-reject safety as "
                 + "insert_field, plus the grid-consistency backstop.")]
    public static ToolResponse SplitCells(
        [Description("Absolute path to the .docx layout file.")] string layoutPath,
        [Description("0-based table index in document order (see get_layout_info's tables[]).")] int tableIndex,
        [Description("0-based row index within the table.")] int row,
        [Description("0-based physical cell index of the spanned cell to split.")] int cellIndex,
        [Description("Which OOXML part the table is in: 'body' (default), 'header', or 'footer'.")] string layoutPart = "body",
        [Description("Only used when layoutPart is 'header'/'footer': a specific part file name; omit for the "
                     + "first.")] string? partName = null)
    {
        return GuardTableEdit(layoutPath, doc =>
        {
            var part = ParseLayoutPartOrThrow(layoutPart);
            return TableStructureEditor.SplitCell(doc, part, partName, tableIndex, row, cellIndex);
        });
    }

    [McpServerTool(Name = "set_cell_borders")]
    [Description("Draw (or clear) the horizontal/vertical RULES on one row of a table - the per-cell "
                 + "w:tcBorders real BC documents get their look from. A BC layout draws no table grid: it "
                 + "has a 1/2-pt rule under the lines-table header row (insert_repeater_table's look='bc' "
                 + "default already adds that one) and a rule above a totals block, and nothing else. This "
                 + "is the tool for the rest of them: 'put a line above the totals row', 'underline the "
                 + "grand-total cell'. Addressed by 0-based table index and row index (the SAME indices "
                 + "get_layout_info reports); omit col to apply to EVERY cell in the row (the usual case - a "
                 + "rule spans the whole row), or pass a 0-based PHYSICAL cell index to target one cell. "
                 + "edges is a comma-separated list of 'top','bottom','left','right' (or 'all'); "
                 + "style='single' (default) draws a line, style='none' clears those edges explicitly (they "
                 + "stay clear even in a table that declares its own border grid). Edges you do NOT name are "
                 + "left exactly as they were. Cosmetic only - no cell, row, column, grid or binding is "
                 + "touched, so unlike the other table tools this one also works on tables that use "
                 + "vMerge. Same write/validate/save-or-reject safety as insert_field.")]
    public static ToolResponse SetCellBorders(
        [Description("Absolute path to the .docx layout file.")] string layoutPath,
        [Description("0-based table index in document order (see get_layout_info's tables[]).")] int tableIndex,
        [Description("0-based row index within the table.")] int row,
        [Description("Comma-separated edges to set: any of 'top','bottom','left','right', or 'all'.")]
        string edges,
        [Description("Optional 0-based PHYSICAL cell index within the row; omit to apply to every cell in "
                     + "the row (the usual case - a rule spans the whole row).")] int? col = null,
        [Description("'single' (default) draws a line of the given size; 'none' explicitly clears the named "
                     + "edges.")] string style = "single",
        [Description("Rule thickness in EIGHTHS of a point (2-96); default 4 = 1/2 pt, the only thickness "
                     + "the real BC corpus uses.")] int size = CellBorderOptions.DefaultSizeEighthPoints,
        [Description("Which OOXML part the table is in: 'body' (default), 'header', or 'footer'.")] string layoutPart = "body",
        [Description("Only used when layoutPart is 'header'/'footer': a specific part file name (e.g. "
                     + "'header2.xml'); omit to target the DEFAULT header/footer part (see get_layout_info's "
                     + "partDetails for each part's role).")] string? partName = null)
    {
        return GuardTableEdit(layoutPath, doc =>
        {
            var part = ParseLayoutPartOrThrow(layoutPart);
            var options = ParseCellBorderOptionsOrThrow(edges, style, size);
            return TableStructureEditor.SetCellBorders(doc, part, partName, tableIndex, row, col, options);
        });
    }

    // ---- flat-argument parsing helpers (private to this family's five table-structure tools plus the
    // repeater-table author tool above) ----

    /// <summary>Splits a comma-separated, agent-supplied list, trimming whitespace and dropping empty entries.</summary>
    private static string[] SplitCommaList(string value) =>
        (value ?? string.Empty).Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

    /// <summary>
    /// Parses <c>insert_repeater_table</c>'s optional comma-separated column-widths string into ints, or
    /// null when not supplied. A malformed entry throws <see cref="ArgumentException"/> (mapped to
    /// <c>invalid_argument</c> by <see cref="ToolGuards.Guard"/>, same as every other bad-input path here); the
    /// actual count-matches-columns check is left to <see cref="LayoutEditor.InsertRepeaterTable"/>/
    /// <see cref="SdtFactory.BuildRepeaterTable"/>, which already validate it against the resolved column list.
    /// </summary>
    private static List<int>? ParseColumnWidths(string? columnWidths)
    {
        if (string.IsNullOrWhiteSpace(columnWidths))
        {
            return null;
        }

        return columnWidths
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(w => int.TryParse(w, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n)
                ? n
                : throw new ArgumentException(
                    $"columnWidths entry '{w}' is not a valid integer.", nameof(columnWidths)))
            .ToList();
    }

    /// <summary>Parses a required comma-separated integer list (e.g. <c>set_column_widths</c>' twips), rejecting an empty list.</summary>
    private static List<int> ParseIntListOrThrow(string value, string paramName)
    {
        var parts = (value ?? string.Empty).Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length == 0)
        {
            throw new ArgumentException(
                $"{paramName} must be a non-empty comma-separated list of integers (twips).", paramName);
        }

        return parts
            .Select(p => int.TryParse(p, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n)
                ? n
                : throw new ArgumentException($"{paramName} entry '{p}' is not a valid integer.", paramName))
            .ToList();
    }

    /// <summary>
    /// Parses a <see cref="LayoutPart"/> from a tool's flat string parameter, throwing <see cref="ArgumentException"/>
    /// on an unknown value. Delegates to <see cref="ToolGuards.TryParseLayoutPart"/> (shared with
    /// <see cref="ToolGuards.BuildLocation"/>) since these five tools address a table directly by
    /// <see cref="LayoutPart"/> rather than building a full <see cref="Location"/>.
    /// </summary>
    private static LayoutPart ParseLayoutPartOrThrow(string layoutPart)
    {
        if (!TryParseLayoutPart(layoutPart, out var part))
        {
            throw new ArgumentException(
                $"Unknown layoutPart '{layoutPart}'. Use 'body', 'header', or 'footer'.", nameof(layoutPart));
        }

        return part;
    }

    /// <summary>Parses a <see cref="TableBorderLook"/> from <c>insert_repeater_table</c>'s <c>look</c> parameter (case-insensitive).</summary>
    private static TableBorderLook ParseTableBorderLookOrThrow(string look) => look.Trim().ToLowerInvariant() switch
    {
        "bc" => TableBorderLook.Bc,
        "grid" => TableBorderLook.Grid,
        _ => throw new ArgumentException(
            $"Unknown look '{look}'. Use 'bc' (default - no drawn grid, just a rule under the header row, "
            + "like every real BC lines table) or 'grid' (an explicit border on every edge).", nameof(look)),
    };

    /// <summary>
    /// Parses <c>set_cell_borders</c>' flat <c>edges</c>/<c>style</c>/<c>size</c> parameters into a
    /// <see cref="CellBorderOptions"/>. The "no edge selected"/"size out of range" rejections belong to
    /// <see cref="TableStructureEditor.SetCellBorders"/> (they are domain rules, checked there for every
    /// caller); this only rejects what the flat string form itself can get wrong — an unknown edge name or
    /// style.
    /// </summary>
    private static CellBorderOptions ParseCellBorderOptionsOrThrow(string edges, string style, int size)
    {
        var remove = style.Trim().ToLowerInvariant() switch
        {
            "single" or "" => false,
            "none" or "nil" => true,
            _ => throw new ArgumentException(
                $"Unknown style '{style}'. Use 'single' (draw a line) or 'none' (clear the named edges).",
                nameof(style)),
        };

        var options = new CellBorderOptions { Remove = remove, SizeEighthPoints = size };
        foreach (var edge in SplitCommaList(edges))
        {
            options = edge.ToLowerInvariant() switch
            {
                "top" => options with { Top = true },
                "bottom" => options with { Bottom = true },
                "left" => options with { Left = true },
                "right" => options with { Right = true },
                "all" => options with { Top = true, Bottom = true, Left = true, Right = true },
                _ => throw new ArgumentException(
                    $"Unknown edge '{edge}'. Use a comma-separated list of 'top', 'bottom', 'left', 'right', "
                    + "or 'all'.", nameof(edges)),
            };
        }

        return options;
    }

    /// <summary>Parses an <see cref="InsertColumnMode"/> from <c>insert_column</c>'s <c>mode</c> parameter (case-insensitive).</summary>
    private static InsertColumnMode ParseInsertColumnModeOrThrow(string mode) => mode.Trim().ToLowerInvariant() switch
    {
        "field" => InsertColumnMode.Field,
        "label" => InsertColumnMode.Label,
        "plaintext" or "plain_text" or "text" => InsertColumnMode.PlainText,
        _ => throw new ArgumentException(
            $"Unknown mode '{mode}'. Use 'field', 'label', or 'plainText'.", nameof(mode)),
    };
}
