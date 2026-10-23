using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;
using ModelContextProtocol.Server;
using static BcWordLayout.McpHost.Tools.ToolGuards;

namespace BcWordLayout.McpHost.Tools;

/// <summary>
/// MCP tools that add or change a table's STRUCTURE (as opposed to <see cref="EditTools"/>, which edits an
/// existing control's or plain-text cell's content): authoring a whole new repeater table
/// (<c>insert_repeater_table</c>) or plain table (<c>insert_table</c>), a nested detail row inside a
/// repeater's item (<c>insert_repeater_row</c>), appending static rows — table-level totals/spacer rows
/// (<c>insert_table_row</c>) and per-group rows inside a repeater's item (<c>insert_subtotal_row</c>) —
/// resizing columns (<c>set_column_widths</c>), adding/removing a column
/// (<c>insert_column</c>/<c>remove_column</c>), horizontally merging/splitting cells within a row
/// (<c>merge_cells</c>/<c>split_cells</c>), and cell rules and fills (<c>set_cell_format</c>).
/// <c>insert_repeater_table</c> sits here rather than in
/// <see cref="EditTools"/> because it is fundamentally a table-shape operation (it builds the table's grid
/// and rows from scratch), even though its edit-safety wrapper is the content-control one
/// (<see cref="ToolGuards.GuardEdit"/>, which <c>insert_repeater_row</c> and <c>insert_subtotal_row</c> —
/// repeater-item edits — also use); the table-structure rest route through
/// <see cref="ToolGuards.GuardTableEdit"/>. Both wrappers, and the grid-consistency backstop
/// <see cref="ToolGuards.GuardMutate{TResult}"/> runs for every mutating tool, live in <see cref="ToolGuards"/>,
/// which also holds the single shared per-path edit lock these tools serialize against alongside every
/// other tool family's mutating tools.
/// </summary>
[McpServerToolType]
public static class TableTools
{
    /// <summary>
    /// Upper bound for any single column width these tools accept, in twips: 22 inches (31680 twips), Word's
    /// maximum page width — no column wider than the widest page Word can lay out can render. The lower bound
    /// is 1: a negative <c>w:gridCol/@w:w</c> is schema-invalid (it surfaced as <c>edit_would_corrupt</c>, a
    /// "bug in the tool" refusal, for what is a mistake in the call) and a zero one writes a column Word
    /// collapses to nothing (GitHub issue #98). Checked when the argument is parsed, naming the argument.
    /// </summary>
    internal const int MaxColumnWidthTwips = 22 * 1440;

    /// <summary>
    /// Upper bound for a cells-DSL <c>N:</c> span: 63, Word's maximum number of columns in a table, so no
    /// cell can span more grid columns than a table can have. The lower bound is 1 — a zero or negative span
    /// used to reach the domain layer and fail there as a sum mismatch or an index-out-of-range crash.
    /// </summary>
    internal const int MaxCellSpan = 63;

    /// <summary>The <c>tableStyle</c> parameter's description, shared by the two table-creating tools.</summary>
    private const string TableStyleDescription =
        "Optional id of a TABLE style the layout's styles part defines, referenced via w:tblStyle - e.g. "
        + "'TableGrid' (create_layout's blank layouts define it) or a branded templatePath shell's own table style. "
        + "Refused when the layout does not define it; the error lists the ones it does. Ids are case-sensitive.";

    [McpServerTool(Name = "insert_repeater_table")]
    [Description("Insert a complete repeater TABLE bound to a repeating data item (e.g. '/Header/Line', as "
                 + "reported by list_dataset_fields): a header row (label controls where a label column "
                 + "exists for a given column - by default suffixed Lbl/_Lbl, see list_dataset_fields's "
                 + "isLabel flag - else static humanized text) plus one data row wrapped in a "
                 + "repeating-section control, with one bound field control per column. This is the flagship "
                 + "tool for adding a new line-items-style table to a layout. By default the table gets the "
                 + "BC-NATIVE look (borders='headerRule'): no drawn grid, just a rule under the header row, exactly "
                 + "like every real BC lines table. TOTALS: the stock shape puts them INSIDE this table as trailing "
                 + "static rows - append them with insert_table_row (+ set_cell_format for the rule). "
                 + "The result reports tableIndex and dataRowIndex - the new table's own 0-based index and "
                 + "the 0-based index of its repeating DATA row - so follow-up edits need no re-read. "
                 + "NESTING: for per-line detail (components, serial/lot nos), the STANDARD BC shape is a "
                 + "detail ROW under the line row - use insert_repeater_row with this call's returned "
                 + "controlId; per-GROUP spacer/subtotal rows are insert_subtotal_row on the same id. "
                 + "Hosting a whole nested TABLE inside a cell (locationType='tableCell' with "
                 + "this call's tableIndex/dataRowIndex) also works but is NOT how standard BC documents "
                 + "are laid out - prefer insert_repeater_row. "
                 + "BODY ONLY: a repeater table in a header or footer is not supported, because nothing shows "
                 + "that Business Central renders a repeating section there (a plain insert_table works in both). "
                 + "Same write/validate/save-or-reject safety as insert_field.")]
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
        [Description(TableStyleDescription)] string? tableStyle = null,
        [Description("Optional comma-separated column widths in twips (1/20 pt, each 1-31680), one per column; count must "
                     + "match the number of columns. Omit for an even split of the document's content width "
                     + "(page width minus margins; for a table placed inside a cell, that cell's width), so "
                     + "the table always fits the page however many columns it has. The widths "
                     + "are BINDING, not advisory: the table is emitted with w:tblLayout='fixed' and a "
                     + "w:tblW total equal to their sum, so Word renders the grid as specified instead of "
                     + "autofitting the columns to whatever the merged content happens to be.")]
        string? columnWidths = null,
        [Description("Optional comma-separated per-column alignments ('left'/'center'/'right', one per "
                     + "column), applied to both the header and data cell of each column - real BC line "
                     + "tables right-align their quantity/price/amount columns. Omit for the default (left).")]
        string? columnAlignments = null,
        [Description("The table's borders, the same option insert_table takes: 'headerRule' (default) draws NO "
                     + "table grid, only a 1/2-pt rule under the header row - the corpus-verified look of every "
                     + "real BC lines table; 'none' draws no border at all (e.g. a composable body whose theme's "
                     + "table style supplies the look); 'grid' draws an explicit single-line border on every edge "
                     + "and between every cell. Use set_cell_format afterwards for the other rules a BC document "
                     + "has (e.g. above a totals block).")] string borders = "headerRule")
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
                Borders = ParseBordersOrThrow(borders, TableBorderLook.HeaderRule),
                HeaderFromLabels = headerFromLabels,
                TableStyle = string.IsNullOrWhiteSpace(tableStyle) ? null : tableStyle,
                ColumnWidths = ParseColumnWidths(columnWidths),
                ColumnAlignments = string.IsNullOrWhiteSpace(columnAlignments) ? null : SplitCommaList(columnAlignments),
            };

            var location = BuildLocation(locationType, controlId, tableIndex, row, col, searchText);
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
                     + "optional 'N:' span prefix (N from 1 to 63). Spans must sum to the parent table's columnCount.")]
        string cells,
        [Description("Optional comma-separated per-cell alignments ('left'/'center'/'right', or '-' to keep "
                     + "the default), one per cell spec.")] string? alignments = null)
    {
        return GuardEdit(layoutPath, doc =>
            LayoutEditor.InsertRepeaterRow(doc, parentControlId, dataItem, ParseRepeaterRowCells(cells, alignments)));
    }

    [McpServerTool(Name = "insert_subtotal_row")]
    [Description("Append a STATIC per-group row at the END of an existing repeater's item - the stock BC "
                 + "shape for GROUP SUBTOTALS (corpus: SalespersonCommission ends each salesperson's "
                 + "repeating item with an empty spacer row and a bold subtotal row, AFTER the nested "
                 + "detail rows). The row renders once per PARENT row (per group), never per detail row - "
                 + "unlike insert_table_row, whose row renders exactly once for the whole table. "
                 + "parentControlId is the repeater's controlId (from insert_repeater_table's result or "
                 + "get_layout_info). cells: same DSL as insert_table_row - '-' = empty spacer cell, a FULL "
                 + "dataset path = a bound cell (a subtotal row binds whatever the dataset provides: the "
                 + "corpus binds a sibling non-repeating Subtotals item's columns and Labels captions; "
                 + "label-shaped paths become label controls), 'a+b' chains paths in one cell, optional "
                 + "'N:' span prefix; spans must sum to the parent table's columnCount. Repeated calls "
                 + "stack rows in order - the stock shape is one all-spacer row (e.g. '9:-') then the bold "
                 + "subtotal row (format {\"bold\": true}, right-aligned amount cells). The GROUP-HEADER half of the "
                 + "stock group shape needs no extra call: the repeater's own line row (built by "
                 + "insert_repeater_table) is the per-group header. The response's controlId is 0 (the row "
                 + "is a plain w:tr; its bound cells' ids are in the summary). Same "
                 + "write/validate/save-or-reject safety as insert_field, plus the grid-consistency "
                 + "backstop.")]
    public static ToolResponse InsertSubtotalRow(
        [Description("Absolute path to the .docx layout file.")] string layoutPath,
        [Description("The repeater's controlId (from insert_repeater_table's result or "
                     + "get_layout_info).")] int parentControlId,
        [Description("Comma-separated cell specs laying the row on the parent table's grid: '-' = empty "
                     + "spacer, a FULL dataset path = one bound cell, 'a+b' = several paths inline in one "
                     + "cell, optional 'N:' span prefix (N from 1 to 63). Spans must sum to the parent table's columnCount.")]
        string cells,
        [Description("Optional comma-separated per-cell alignments ('left'/'center'/'right', or '-' to keep "
                     + "the default), one per cell spec - subtotal amounts are right-aligned in the corpus.")]
        string? alignments = null,
        [Description(TextFormatArgument.RowDescription)] TextFormatArgument? format = null)
    {
        return GuardEdit(layoutPath, doc =>
        {
            return LayoutEditor.InsertSubtotalRow(
                doc, parentControlId, ParseRepeaterRowCells(cells, alignments), TextFormatArgument.ToDomain(format));
        });
    }

    /// <summary>
    /// Parses the cells DSL (<c>[N:]entry(+entry)*</c> or <c>-</c>, comma-separated) plus the optional
    /// per-cell alignment list into <see cref="RepeaterRowCell"/>s — shared by <c>insert_repeater_row</c>
    /// (entries are leaf column names of the child data item) and <c>insert_table_row</c> (entries are
    /// FULL dataset paths); the syntax is identical, only what an entry means differs, and that is the
    /// domain layer's to validate. Malformed input throws <see cref="ArgumentException"/> (mapped to
    /// <c>invalid_argument</c>); semantic validation (columns/paths exist, spans cover the grid) is the
    /// domain layer's.
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

                if (span is < 1 or > MaxCellSpan)
                {
                    throw new ArgumentException(
                        $"cells entry '{spec}' spans {span} grid column(s); a span must be between 1 and "
                        + $"{MaxCellSpan} (omit the 'N:' prefix for a single column).",
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

    [McpServerTool(Name = "insert_table_row")]
    [Description("Insert ONE STATIC (non-repeating) row into an EXISTING table - the stock BC shape for "
                 + "the totals block INSIDE the line-items table: every stock document layout ends its "
                 + "lines table with right-anchored totals rows (leading empty spacer cells + a gridSpan-"
                 + "merged caption cell + the amount cell, bold, with a rule) rather than a separate table. "
                 + "cells lays the row out on the table's grid, comma-separated, one entry per cell: '-' is "
                 + "an empty spacer cell, a FULL dataset path (e.g. '/Header/Totals/TotalAmountIncludingVAT' "
                 + "- unlike insert_repeater_row's leaf names, because a totals row binds whatever the "
                 + "dataset provides, typically a NON-repeating Totals item; label-shaped paths become "
                 + "label controls) is a bound cell, 'a+b' chains several paths inline in ONE cell, and an "
                 + "optional 'N:' prefix makes the cell span N grid columns (e.g. "
                 + "'4:-,3:/Header/Totals/TotalIncludingVATText,/Header/Totals/TotalAmountIncludingVAT'). "
                 + "The spans MUST sum to the table's columnCount. atRow is the 0-based row position to "
                 + "insert at, in get_layout_info's tables[] numbering (one row per w:tr, a repeater's rows "
                 + "included); a position BETWEEN two rows of the same repeater is refused, since a static "
                 + "row cannot sit inside a repeating item - use the index of the repeater's first row "
                 + "(before it) or of the row after its last (after it). Omit atRow to append after the "
                 + "last row, where a totals row goes. The row renders exactly ONCE - it is never part of a "
                 + "repeating section (for "
                 + "per-group subtotal rows inside a repeater's item, use insert_subtotal_row). Use format {\"bold\": true} "
                 + "for the stock grand-total look and alignments for right-anchored amounts, then "
                 + "set_cell_format for the rule and set_cell_text for any literal caption. REJECTED for "
                 + "tables using w:vMerge. Same write/validate/save-or-reject safety as insert_field, plus "
                 + "the grid-consistency backstop.")]
    public static ToolResponse InsertTableRow(
        [Description("Absolute path to the .docx layout file.")] string layoutPath,
        [Description("0-based table index in document order (see get_layout_info's tables[]).")] int tableIndex,
        [Description("Comma-separated cell specs laying the row on the table's grid: '-' = empty spacer, "
                     + "a FULL dataset path = one bound cell, 'a+b' = several paths inline in one cell, "
                     + "optional 'N:' span prefix (N from 1 to 63). Spans must sum to the table's columnCount.")] string cells,
        [Description("Optional 0-based row position (0..rowCount), not inside a repeater (see this tool's "
                     + "description). Omit, or pass the table's rowCount, to append after the last row - where "
                     + "a totals row goes.")] int? atRow = null,
        [Description("Optional comma-separated per-cell alignments ('left'/'center'/'right', or '-' to keep "
                     + "the default), one per cell spec - amounts in a BC totals row are right-aligned.")]
        string? alignments = null,
        [Description(TextFormatArgument.RowDescription)] TextFormatArgument? format = null,
        [Description("Which OOXML part the table is in: 'body' (default), 'header', or 'footer'.")] string layoutPart = "body",
        [Description("Only used when layoutPart is 'header'/'footer': a specific part file name (e.g. "
                     + "'header2.xml'); omit to target the DEFAULT header/footer part (see get_layout_info's "
                     + "partDetails for each part's role).")] string? partName = null)
    {
        return GuardTableEdit(layoutPath, doc =>
        {
            var part = ParseLayoutPartOrThrow(layoutPart);
            var cellList = ParseRepeaterRowCells(cells, alignments);
            return TableStructureEditor.InsertStaticRow(doc, part, partName, tableIndex, atRow, cellList, TextFormatArgument.ToDomain(format));
        });
    }

    [McpServerTool(Name = "insert_table")]
    [Description("Insert a NEW plain (unbound) table - the building block for every NON-repeating layout "
                 + "section a real BC document is made of: side-by-side address columns, label/value "
                 + "header-info grids, and a stand-alone right-anchored totals block. (For the LINE-ITEMS "
                 + "table bound to a repeating data item use insert_repeater_table instead, and for the "
                 + "stock totals-INSIDE-the-lines-table shape use insert_table_row on that table.) The table is rows x columns of "
                 + "empty single-column cells, borderless by default (the corpus shape for those blocks); "
                 + "borders and tableStyle work exactly as on insert_repeater_table. "
                 + "columnWidths (twips, comma-separated, one per column) defaults to an even split of the "
                 + "layout's own content width (its page width minus the left and right margins, or 10206 - "
                 + "the A4 content width - when the layout declares no page size or margins); "
                 + "columnAlignments (left/center/right, comma-separated, one "
                 + "per column) seeds each cell's paragraph justification - set_cell_text and "
                 + "insert_field/insert_label (locationType 'tableCell') preserve it afterwards. The result's "
                 + "tableIndex is the new table's 0-based index within its part: EXACTLY the index to pass, "
                 + "with the same layoutPart/partName, to set_cell_text/insert_field/insert_label/"
                 + "set_column_widths next to fill the cells. Works in the body, a header or a footer, like "
                 + "insert_field: every stock header and footer is built from such tables (the logo beside "
                 + "the document details, the company details across the footer). Same "
                 + "write/validate/save-or-reject safety as insert_field.")]
    public static ToolResponse InsertTable(
        [Description("Absolute path to the .docx layout file.")] string layoutPath,
        [Description("Number of rows (1-100).")] int rows,
        [Description("Number of columns (1-30) - a count, unlike insert_repeater_table's list of column names.")]
        int columns,
        [Description("Where to insert: 'documentEnd', 'afterControl', 'tableCell', or 'atText' "
                     + "(case-insensitive).")] string locationType,
        [Description("Required for 'afterControl': the w:id of the control to insert after.")] int? controlId = null,
        [Description("Required for 'tableCell': 0-based table index in document order.")] int? tableIndex = null,
        [Description("Required for 'tableCell': 0-based row index within the table.")] int? row = null,
        [Description("Required for 'tableCell': 0-based cell/column index within the row.")] int? col = null,
        [Description("Required for 'atText': substring to search for in existing run text (ordinal match).")]
        string? searchText = null,
        [Description("Optional comma-separated column widths in twips (each 1-31680), one per column; omit for an even "
                     + "split of the layout's content width (page width minus margins; 10206 twips when the "
                     + "layout declares neither). The widths are BINDING, not "
                     + "advisory: the table is emitted with w:tblLayout='fixed' and a w:tblW total equal to "
                     + "their sum, so Word renders the grid as specified instead of autofitting the columns "
                     + "to their content.")] string? columnWidths = null,
        [Description("Optional comma-separated per-column alignments ('left'/'center'/'right', one per "
                     + "column) seeding each cell's paragraph justification; omit for the default (left).")]
        string? columnAlignments = null,
        [Description("The table's borders, the same option insert_repeater_table takes: 'none' (default) is "
                     + "borderless - the shape of BC's own address/info/totals blocks; 'headerRule' draws a 1/2-pt "
                     + "rule under the first row, the rule set_cell_format draws for a bottom edge; 'grid' draws an "
                     + "explicit single-line border on every edge and between every cell.")]
        string borders = "none",
        [Description(TableStyleDescription)] string? tableStyle = null,
        [Description("Which OOXML part locationType resolves within: 'body' (default), 'header', or "
                     + "'footer' (case-insensitive).")] string layoutPart = "body",
        [Description("Only used when layoutPart is 'header'/'footer': a specific part file name, e.g. "
                     + "'header2.xml' (see get_layout_info's partDetails for the names actually present). Omit "
                     + "to target the DEFAULT part - the everyday header/footer, which is often NOT the first "
                     + "part in the package (partDetails gives each part's role).")] string? partName = null)
    {
        return GuardTableEdit(layoutPath, doc =>
        {
            var location = BuildLocation(locationType, controlId, tableIndex, row, col, searchText, layoutPart, partName);
            var alignments = string.IsNullOrWhiteSpace(columnAlignments)
                ? null
                : SplitCommaList(columnAlignments);
            return LayoutEditor.InsertPlainTable(
                doc, rows, columns, location, ParseColumnWidths(columnWidths), alignments,
                ParseBordersOrThrow(borders, TableBorderLook.None),
                string.IsNullOrWhiteSpace(tableStyle) ? null : tableStyle);
        }, tool: "insert_table"); // rows/columns mean something else on other tools - see ToolSpecificInvalidArgumentHint
    }

    [McpServerTool(Name = "set_column_widths")]
    [Description("Set the column widths of a table, addressed by its 0-based table index (the SAME index "
                 + "get_layout_info reports in its tables[] section). widths is an array of twips (1/20 pt), "
                 + "ONE PER GRID COLUMN - the count must equal the table's columnCount from "
                 + "get_layout_info, and a cell's own columns are list positions gridColumn..gridColumn+"
                 + "gridSpan-1 (get_layout_info's tables[].rows[].cells[] report both; colIndex is the "
                 + "physical position in the row, not a grid column). Each grid column's w:gridCol width is "
                 + "set, and every cell's width (w:tcW) "
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
        [Description("The column widths in twips, each 1 to 31680, exactly one per GRID column - the count must "
                     + "match the table's columnCount - e.g. [1500, 6206, 2500].")]
        [MinLength(1)]
        [Range(1, MaxColumnWidthTwips)]
        IReadOnlyList<int> widths,
        [Description("Which OOXML part the table is in: 'body' (default), 'header', or 'footer'.")] string layoutPart = "body",
        [Description("Only used when layoutPart is 'header'/'footer': a specific part file name (e.g. "
                     + "'header2.xml'); omit to target the DEFAULT header/footer part (see get_layout_info's "
                     + "partDetails for each part's role).")] string? partName = null)
    {
        return GuardTableEdit(layoutPath, doc =>
        {
            var part = ParseLayoutPartOrThrow(layoutPart);
            return TableStructureEditor.SetColumnWidths(doc, part, partName, tableIndex, CheckWidthsOrThrow(widths, nameof(widths)));
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
                 + "between existing columns - a cell's gridColumn in get_layout_info's tables[] puts the new "
                 + "column just before that cell, its gridColumn+gridSpan just after it (colIndex is the "
                 + "physical position in the row, not a grid position). An interior position must land on a "
                 + "cell BOUNDARY in every row "
                 + "that carries the new column's content (a header row, and the bound data row); a row with "
                 + "nothing there - a spacer, or a totals block's leading empty run - has its spanning cell "
                 + "widened instead of gaining a cell, so the block stays put. If it falls INSIDE a spanned "
                 + "cell of a content row the call is refused, naming the cell, rather than guessing (use "
                 + "split_cells there first). mode selects what the new column holds: 'field' or "
                 + "'label' bind the data cell via dataPath (a full dataset path exactly like insert_field/"
                 + "insert_label, e.g. '/Header/Line/Discount_Line' - it may bind a column of the repeater's "
                 + "data item OR a parent/Header field); 'plainText' adds an unbound column (headerText plus "
                 + "empty cells). In a line-items table the bound control goes ONLY into the rows of the "
                 + "repeater whose data item contains dataPath (a /Header/Line field into the Line row, never "
                 + "into a ReportTotalsLine row of the same table, which is treated like any other non-content "
                 + "row) - and into ONE row per repeater item: an item with several rows (spacer, field, "
                 + "nested detail, subtotal rows) gets it in the first row that already holds a field of "
                 + "dataPath's data item, else its first row; every other row gets filler. A parent/Header "
                 + "field is accepted only when the table has a single repeater (detail repeaters nested in "
                 + "it do not count), and a "
                 + "dataPath matching none of the table's repeaters is refused. In a table with no repeater it "
                 + "goes into every non-header row. Header rows get a header cell (bind headerLabelPath, else "
                 + "headerText, else a humanized column name): rows marked w:tblHeader, else the first row of "
                 + "a line-items table, else - in a plain table - its caption rows (rows whose bound controls "
                 + "are all labels, the BC caption-over-value shape). headerText/headerLabelPath on a table "
                 + "with no header row are refused rather than dropped, and headerLabelPath is checked against "
                 + "the dataset in every mode. All other rows get an empty cell. One w:gridCol is added and "
                 + "the table KEEPS its total width, like remove_column: the new column's width is taken "
                 + "proportionally from the existing columns and every row's cell widths are re-sliced from "
                 + "the new grid (cells declared in percent become twips), so a full-width table stays "
                 + "within the margins and a percentage-width table does not squeeze the new column. "
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
                     + "'plainText'). Omit to humanize the dataPath leaf name. Refused for a table with no "
                     + "header row (see this tool's description).")] string? headerText = null,
        [Description("Optional: bind the header cell to this label-column dataset path (by default a "
                     + "'*Lbl'/'*_Lbl' path - see list_dataset_fields's isLabel flag) instead of static text.")]
        string? headerLabelPath = null,
        [Description("Optional 0-based GRID position for the new column (0..columnCount) - a cell's gridColumn "
                     + "(before it) or gridColumn+gridSpan (after it) from get_layout_info's tables[], never "
                     + "its colIndex. Omit, or pass the table's columnCount, to append at the far-right edge; "
                     + "an interior value inserts between existing columns (see this tool's description for "
                     + "the spanned-cell rule).")]
        int? atColumn = null,
        [Description("Optional width of the NEW column in twips, taken proportionally from the existing "
                     + "columns (the table's total width does not change), so it must be positive and less "
                     + "than the table's current total. Omit for an equal share: total / (columnCount + 1). "
                     + "Follow with set_column_widths for any other distribution.")]
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
                // The upper bound only bites where the domain cannot check against the table's own total
                // (a grid whose widths do not parse); "smaller than that total" stays the domain's rule.
                Width = width is { } w ? CheckColumnWidth(w, $"width {w}", nameof(width)) : null,
            };
            return TableStructureEditor.InsertColumn(doc, part, partName, tableIndex, atColumn, options);
        });
    }

    [McpServerTool(Name = "remove_column")]
    [Description("Remove a whole column from a table, addressed by its 0-based table index and 0-based GRID "
                 + "column index: the gridColumn get_layout_info's tables[].rows[].cells[] report for the "
                 + "column's cells - NOT their colIndex, which is the cell's physical position in its row and "
                 + "names a different column on any row with spanned cells (pervasive in real BC tables). The "
                 + "summary says, per row, whether a cell was deleted, a spanned cell narrowed, or a skipped "
                 + "run shrunk, and warns when no row had a cell of its own there. The matching w:gridCol is "
                 + "removed, and in "
                 + "every row the cell covering that grid column is deleted when it spans only that column - "
                 + "including a bound field/label cell, whose binding is dropped with it (UNLIKE remove_control, "
                 + "which preserves the cell) - or has its w:gridSpan decremented when it spans more. Repeater "
                 + "and cell-level wrappers are preserved (only the inner cells/rows change). The last "
                 + "remaining column cannot be removed, and neither can a column that is some row's ONLY cell "
                 + "(a row otherwise made of w:gridBefore/w:gridAfter skipped columns - Word cannot open a row "
                 + "with no cells); a table none of whose rows matches its w:tblGrid is refused too (the "
                 + "grid itself is out of step). The removed column's width is redistributed "
                 + "proportionally across the remaining columns, so a table that spanned the full content "
                 + "width still does (follow up with set_column_widths if you want a different distribution). "
                 + "REJECTED for tables using w:vMerge (w:gridBefore/"
                 + "w:gridAfter rows are supported). This is the tool for requests like 'remove the GST Amount column from the "
                 + "lines'. Same write/validate/save-or-reject safety as insert_field, plus the "
                 + "grid-consistency backstop that rejects any edit desyncing a row from the grid.")]
    public static ToolResponse RemoveColumn(
        [Description("Absolute path to the .docx layout file.")] string layoutPath,
        [Description("0-based table index in document order (see get_layout_info's tables[]).")] int tableIndex,
        [Description("0-based GRID column index to remove (0..columnCount-1): a cell's gridColumn from "
                     + "get_layout_info's tables[], not its colIndex.")] int column,
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
                 + "(inclusive) within that row - the SAME table/row indices and cell colIndex get_layout_info "
                 + "reports (the result's columnIndex is the merged cell's GRID column). The "
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
                     + "DEFAULT part - the everyday header/footer, which is often NOT the first part in the "
                     + "package (get_layout_info's partDetails gives each part's role).")] string? partName = null)
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
                 + "(the SAME table/row indices and cell colIndex get_layout_info reports). The cell's "
                 + "w:gridSpan is reset to 1 (keeping "
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
                     + "DEFAULT part - the everyday header/footer, which is often NOT the first part in the "
                     + "package (get_layout_info's partDetails gives each part's role).")] string? partName = null)
    {
        return GuardTableEdit(layoutPath, doc =>
        {
            var part = ParseLayoutPartOrThrow(layoutPart);
            return TableStructureEditor.SplitCell(doc, part, partName, tableIndex, row, cellIndex);
        });
    }

    [McpServerTool(Name = "set_cell_format")]
    [Description("Set the APPEARANCE of the cells on one row of a table: the per-cell RULES (w:tcBorders) "
                 + "real BC documents get their look from, and/or a FILL colour (w:shd). A BC layout draws no "
                 + "table grid: it has a 1/2-pt rule under the lines-table header row (insert_repeater_table's "
                 + "borders='headerRule' default already adds that one) and a rule above a totals block, and nothing "
                 + "else. This is the tool for the rest of them - 'put a line above the totals row', "
                 + "'underline the grand-total cell' - and for a shaded header row. Addressed by 0-based table "
                 + "index and row index (the SAME indices get_layout_info reports); omit col to apply to EVERY "
                 + "cell in the row (the usual case), or pass a 0-based PHYSICAL cell index (the cell's "
                 + "colIndex) to target one cell - the result's columnIndex is then that cell's GRID column. "
                 + "edges is an array of 'top', 'bottom', 'left', 'right' (or 'all'); "
                 + "borderStyle='single' (default) draws a line, 'none' clears those edges explicitly (they "
                 + "stay clear even in a table that declares its own border grid). Edges you do NOT name are "
                 + "left exactly as they were. fill is six hex digits RRGGBB (e.g. 'E8E8E8', the stock light "
                 + "grey) or 'auto' for no fill. Pass edges, fill, or both. Cosmetic only - no cell, row, "
                 + "column, grid or binding is touched, so unlike the other table tools this one also works on "
                 + "tables that use vMerge. To format the TEXT in the cells, use set_text_format. Same "
                 + "write/validate/save-or-reject safety as insert_field.")]
    public static ToolResponse SetCellFormat(
        [Description("Absolute path to the .docx layout file.")] string layoutPath,
        [Description("0-based table index in document order (see get_layout_info's tables[]).")] int tableIndex,
        [Description("0-based row index within the table.")] int row,
        [Description("Optional 0-based PHYSICAL cell index within the row; omit to apply to every cell in "
                     + "the row (the usual case - a rule spans the whole row).")] int? col = null,
        [Description("The rules to draw or clear: 'top', 'bottom', 'left', 'right' or 'all', e.g. [\"top\"] for "
                     + "the rule above a totals row. Omit to leave the rules alone.")]
        [MinLength(1)]
        [AllowedValues("top", "bottom", "left", "right", "all")]
        IReadOnlyList<string>? edges = null,
        [Description("For edges: 'single' (the default) draws a line of borderSize; 'none' explicitly clears "
                     + "the named edges.")] string? borderStyle = null,
        [Description("For edges: rule thickness in EIGHTHS of a point (2-96); default 4 = 1/2 pt, the only "
                     + "thickness the real BC corpus uses.")] int? borderSize = null,
        [Description("Cell fill: six hex digits RRGGBB (e.g. 'E8E8E8'; a leading '#' is fine) or 'auto' for no "
                     + "fill. Omit to leave the fill alone.")] string? fill = null,
        [Description("Which OOXML part the table is in: 'body' (default), 'header', or 'footer'.")] string layoutPart = "body",
        [Description("Only used when layoutPart is 'header'/'footer': a specific part file name (e.g. "
                     + "'header2.xml'); omit to target the DEFAULT header/footer part (see get_layout_info's "
                     + "partDetails for each part's role).")] string? partName = null)
    {
        return GuardTableEdit(layoutPath, doc =>
        {
            var part = ParseLayoutPartOrThrow(layoutPart);
            var borders = ParseCellBorderOptionsOrThrow(edges, borderStyle, borderSize);
            return TableStructureEditor.SetCellFormat(doc, part, partName, tableIndex, row, col, borders, fill);
        });
    }

    // ---- flat-argument parsing helpers (private to this family's seven table-structure tools plus the
    // table and repeater authoring tools above) ----

    /// <summary>Splits a comma-separated, agent-supplied list, trimming whitespace and dropping empty entries.</summary>
    private static string[] SplitCommaList(string? value) =>
        (value ?? string.Empty).Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

    /// <summary>
    /// Parses <c>insert_repeater_table</c>'s/<c>insert_table</c>'s optional comma-separated column-widths string
    /// into ints, or null when not supplied. A malformed or out-of-range entry (see
    /// <see cref="MaxColumnWidthTwips"/>) throws <see cref="ArgumentException"/> (mapped to
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
                ? CheckColumnWidth(n, $"columnWidths entry '{w}'", nameof(columnWidths))
                : throw new ArgumentException(
                    $"columnWidths entry '{w}' is not a valid integer.", nameof(columnWidths)))
            .ToList();
    }

    /// <summary>
    /// Checks a required array of column widths in twips (<c>set_column_widths</c>' widths): refuses a missing or
    /// empty one, and an entry outside 1..<see cref="MaxColumnWidthTwips"/>, naming the entry by its position.
    /// The binder has already refused anything that is not an integer.
    /// </summary>
    private static List<int> CheckWidthsOrThrow(IReadOnlyList<int>? widths, string paramName)
    {
        if (widths is null || widths.Count == 0)
        {
            throw new ArgumentException(
                $"{paramName} must list one width in twips per grid column, e.g. [1500, 6206, 2500].", paramName);
        }

        return widths.Select((w, i) => CheckColumnWidth(w, $"{paramName}[{i}] ({w})", paramName)).ToList();
    }

    /// <summary>
    /// Returns <paramref name="twips"/> when it is a usable column width (1..<see cref="MaxColumnWidthTwips"/>),
    /// else throws an <see cref="ArgumentException"/> naming <paramref name="what"/> and keyed on
    /// <paramref name="paramName"/>, so the refusal carries that argument's own hint.
    /// </summary>
    private static int CheckColumnWidth(int twips, string what, string paramName) =>
        twips is >= 1 and <= MaxColumnWidthTwips
            ? twips
            : throw new ArgumentException(
                $"{what} is out of range: a column width must be between 1 and {MaxColumnWidthTwips} twips "
                + "(22 in, Word's maximum page width).",
                paramName);

    /// <summary>
    /// Parses a <see cref="LayoutPart"/> from a tool's flat string parameter, throwing <see cref="ArgumentException"/>
    /// on an unknown value. Delegates to <see cref="ToolGuards.TryParseLayoutPart"/> (shared with
    /// <see cref="ToolGuards.BuildLocation"/>) since these seven tools address a table directly by
    /// <see cref="LayoutPart"/> rather than building a full <see cref="Location"/>. Null is the default
    /// (<c>body</c>) — see <see cref="ToolGuards.TryParseLayoutPart"/>.
    /// </summary>
    private static LayoutPart ParseLayoutPartOrThrow(string? layoutPart)
    {
        if (!TryParseLayoutPart(layoutPart, out var part))
        {
            throw new ArgumentException(
                $"Unknown layoutPart '{layoutPart}'. Use 'body', 'header', or 'footer'.", nameof(layoutPart));
        }

        return part;
    }

    /// <summary>
    /// Parses the <c>borders</c> parameter <c>insert_table</c> and <c>insert_repeater_table</c> share
    /// (case-insensitive). An explicit JSON <c>null</c> (bound straight through by the SDK) is the calling tool's
    /// documented default, <paramref name="whenNull"/>, as if the argument had been omitted.
    /// </summary>
    private static TableBorderLook ParseBordersOrThrow(string? borders, TableBorderLook whenNull) =>
        borders?.Trim().ToLowerInvariant() switch
        {
            null => whenNull,
            "none" => TableBorderLook.None,
            "headerrule" => TableBorderLook.HeaderRule,
            "grid" => TableBorderLook.Grid,
            _ => throw new ArgumentException(
                $"Unknown borders '{borders}'. Use 'none' (no border at all), 'headerRule' (a 1/2-pt rule under the "
                + "first row, like every real BC lines table) or 'grid' (an explicit border on every edge).",
                nameof(borders)),
        };

    /// <summary>
    /// Parses <c>set_cell_format</c>'s flat <c>edges</c>/<c>borderStyle</c>/<c>borderSize</c> parameters into a
    /// <see cref="CellBorderOptions"/>, or null when none of them is given (the call only sets a fill). The "no
    /// edge selected"/"size out of range" rejections belong to <see cref="TableStructureEditor.SetCellFormat"/>
    /// (they are domain rules, checked there for every caller); this rejects what the flat form itself can get
    /// wrong — an unknown edge name or style, and a style or size without edges, which would otherwise be
    /// silently ignored. A null style or size is its documented default (<c>single</c>, ½ pt).
    /// </summary>
    private static CellBorderOptions? ParseCellBorderOptionsOrThrow(IReadOnlyList<string?>? edges, string? borderStyle, int? borderSize)
    {
        if (edges is null)
        {
            if (borderStyle is not null || borderSize is not null)
            {
                throw new ArgumentException(
                    "borderStyle and borderSize describe the edges you name, but edges was not given; pass edges "
                    + "too (e.g. 'top'), or drop them to set only the fill.",
                    nameof(edges));
            }

            return null;
        }

        var remove = (borderStyle ?? "single").Trim().ToLowerInvariant() switch
        {
            "single" or "" => false,
            "none" or "nil" => true,
            _ => throw new ArgumentException(
                $"Unknown borderStyle '{borderStyle}'. Use 'single' (draw a line) or 'none' (clear the named edges).",
                nameof(borderStyle)),
        };

        var options = new CellBorderOptions
        {
            Remove = remove,
            SizeEighthPoints = borderSize ?? CellBorderOptions.DefaultSizeEighthPoints,
        };
        for (var i = 0; i < edges.Count; i++)
        {
            options = edges[i]?.Trim().ToLowerInvariant() switch
            {
                "top" => options with { Top = true },
                "bottom" => options with { Bottom = true },
                "left" => options with { Left = true },
                "right" => options with { Right = true },
                "all" => options with { Top = true, Bottom = true, Left = true, Right = true },
                _ => throw new ArgumentException(
                    $"Unknown edge {(edges[i] is null ? "null" : $"'{edges[i]}'")} at edges[{i}]. Each edge is 'top', "
                    + "'bottom', 'left', 'right' or 'all', e.g. [\"top\", \"bottom\"].", nameof(edges)),
            };
        }

        return options;
    }

    /// <summary>
    /// Parses an <see cref="InsertColumnMode"/> from <c>insert_column</c>'s <c>mode</c> parameter
    /// (case-insensitive). <c>mode</c> is required with no default, so an explicit JSON <c>null</c> is
    /// refused as <c>invalid_argument</c> rather than guessed.
    /// </summary>
    private static InsertColumnMode ParseInsertColumnModeOrThrow(string? mode) => mode?.Trim().ToLowerInvariant() switch
    {
        "field" => InsertColumnMode.Field,
        "label" => InsertColumnMode.Label,
        "plaintext" or "plain_text" or "text" => InsertColumnMode.PlainText,
        _ => throw new ArgumentException(
            (mode is null ? "mode is required (got null)." : $"Unknown mode '{mode}'.")
            + " Use 'field', 'label', or 'plainText'.", nameof(mode)),
    };
}
