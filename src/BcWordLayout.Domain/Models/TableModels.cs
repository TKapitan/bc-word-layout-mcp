namespace BcWordLayout.Domain.Models;

/// <summary>
/// The structural "level" of a content control — i.e. which concrete <c>w:sdt</c> kind it is, which
/// dictates what a whole-control removal would take with it. This is the single most important signal a
/// caller needs BEFORE removing a control: a <see cref="Cell"/> or <see cref="Row"/> control wraps a
/// structural table part (a <c>w:tc</c> cell / <c>w:tr</c> row) that defines the table grid, so deleting
/// the whole control would delete that cell/row and can leave the table with fewer cells than its grid
/// defines (a visually broken table). See <c>LayoutEditor.RemoveControl</c>, which never deletes a
/// cell-level control's <c>w:tc</c> precisely because of this.
/// </summary>
public enum SdtLevel
{
    /// <summary>Inline run-level control (<c>w:sdt</c> as <c>SdtRun</c>) — sits in a paragraph's run flow.</summary>
    Run,

    /// <summary>Block-level control (<c>SdtBlock</c>) — a direct child of a body/cell/etc., wraps block content.</summary>
    Block,

    /// <summary>Cell-level control (<c>SdtCell</c>) — sits where a <c>w:tc</c> would and wraps a whole table cell.</summary>
    Cell,

    /// <summary>Row-level control (<c>SdtRow</c>) — sits where a <c>w:tr</c> would and wraps a whole table row (e.g. a repeater).</summary>
    Row,

    /// <summary>East-Asian ruby (phonetic-guide) run control (<c>SdtRunRuby</c>).</summary>
    RunRuby,

    /// <summary>An <see cref="DocumentFormat.OpenXml.Wordprocessing.SdtElement"/> subclass not in the five above (not expected in practice).</summary>
    Unknown,
}

/// <summary>
/// A single table found in a layout part, described structurally so callers can reason about the grid
/// BEFORE editing it. <see cref="TableIndex"/> is the SAME 0-based, per-part, document-order index the
/// insert tools' <c>tableCell</c> addressing uses (it is derived from <c>root.Descendants&lt;Table&gt;()</c>,
/// exactly like <c>LocationResolver</c>), so a table reported here can be targeted by that index directly.
/// Nested tables therefore each get their OWN top-level index (they are flattened into document order),
/// matching the insert tools.
/// </summary>
public sealed class TableStructure
{
    /// <summary>Which OOXML part the table lives in (e.g. <c>document.xml</c>, <c>header1.xml</c>).</summary>
    public required string Part { get; init; }

    /// <summary>0-based, per-part, document-order table index — the same index <c>tableCell</c> addressing uses.</summary>
    public required int TableIndex { get; init; }

    /// <summary>Number of rows (plain <c>w:tr</c> plus any row-level control rows).</summary>
    public required int RowCount { get; init; }

    /// <summary>Number of grid columns declared by the table's <c>w:tblGrid</c>.</summary>
    public required int ColumnCount { get; init; }

    /// <summary>Declared column widths in twips (from each <c>w:gridCol/@w:w</c>); empty when none are present.</summary>
    public required IReadOnlyList<int> GridColumnWidths { get; init; }

    /// <summary>The rows in document order.</summary>
    public required IReadOnlyList<TableRowInfo> Rows { get; init; }
}

/// <summary>One row of a <see cref="TableStructure"/>.</summary>
public sealed class TableRowInfo
{
    /// <summary>0-based row index within its table.</summary>
    public required int RowIndex { get; init; }

    /// <summary>
    /// True when the row itself is wrapped by a row-level content control (<c>SdtRow</c>) — e.g. a repeater's
    /// <c>w15:repeatingSection</c>/<c>w15:repeatingSectionItem</c> row. Removing such a control (whole-control)
    /// removes the entire row.
    /// </summary>
    public required bool IsControlRow { get; init; }

    /// <summary>The row-level control's <c>w:id</c>, when <see cref="IsControlRow"/> is true; else null.</summary>
    public int? ControlId { get; init; }

    /// <summary>
    /// Grid columns the row skips BEFORE its first cell (<c>w:trPr/w:gridBefore</c>, 0 when absent) — so its
    /// first cell starts at this grid column, not 0. With <see cref="GridAfter"/> and the cells'
    /// <see cref="TableCellInfo.GridSpan"/>s it accounts for the whole grid:
    /// <c>GridBefore + ΣGridSpan + GridAfter = ColumnCount</c> on a well-formed row.
    /// </summary>
    public required int GridBefore { get; init; }

    /// <summary>Grid columns the row skips AFTER its last cell (<c>w:trPr/w:gridAfter</c>, 0 when absent).</summary>
    public required int GridAfter { get; init; }

    /// <summary>The cells in document order.</summary>
    public required IReadOnlyList<TableCellInfo> Cells { get; init; }
}

/// <summary>One cell of a <see cref="TableRowInfo"/>.</summary>
/// <remarks>
/// TWO COORDINATES, ONE PER KIND OF TOOL. <see cref="ColIndex"/> is the cell's PHYSICAL position in its row
/// — what the cell tools take (<c>tableCell</c> locations, <c>set_cell_text</c>, <c>merge_cells</c>,
/// <c>split_cells</c>, <c>set_cell_borders</c>'s <c>col</c>). <see cref="GridColumn"/>/<see cref="GridSpan"/>
/// are where the cell sits on the table's <c>w:tblGrid</c> — what the column tools take
/// (<c>insert_column</c>'s <c>atColumn</c>, <c>remove_column</c>, <c>set_column_widths</c>' list order).
/// They differ on almost every real BC table, which is <c>gridSpan</c>-pervasive: on the stock VAT-spec
/// invoice's lines table, Line Amount is physical cell 7 in both its header and data rows but grid column 9,
/// and <c>remove_column</c> with the physical index removed a grid column inside the VAT % cells instead
/// (GitHub issue #81). Before <see cref="GridColumn"/> was reported, nothing in <c>get_layout_info</c> let a
/// caller convert one to the other.
/// </remarks>
public sealed class TableCellInfo
{
    /// <summary>0-based column index within its row (counts every cell, whether or not it is a control cell).</summary>
    public required int ColIndex { get; init; }

    /// <summary>
    /// 0-based GRID column the cell starts at — its row's <see cref="TableRowInfo.GridBefore"/> plus the
    /// <see cref="GridSpan"/>s of the cells before it. The index to pass to the column tools.
    /// </summary>
    public required int GridColumn { get; init; }

    /// <summary>How many grid columns the cell covers (<c>w:gridSpan</c>, 1 when absent): grid columns <see cref="GridColumn"/>..<see cref="GridColumn"/>+<see cref="GridSpan"/>-1.</summary>
    public required int GridSpan { get; init; }

    /// <summary>
    /// True when the cell itself is wrapped by a cell-level content control (<c>SdtCell</c>) — e.g. the BC
    /// address fields <c>CustomerAddress1..6</c>. Removing such a control WITHOUT keeping the column would,
    /// naively, delete the whole <c>w:tc</c> and break the grid; <c>remove_control</c> guards against that.
    /// </summary>
    public required bool IsControlCell { get; init; }

    /// <summary>
    /// The cell-level control's <c>w:id</c>, when <see cref="IsControlCell"/> is true; else null. When the cell
    /// is wrapped in a CHAIN of cell-level controls (the Business Central Word add-in's "Hide Table Row If
    /// Empty" wrapper around a bound field's own cell-level control), this — and <see cref="ControlKind"/>/
    /// <see cref="Alias"/>/<see cref="XPath"/> — describe the INNERMOST one, the control directly around the
    /// <c>w:tc</c>; the outer wrappers are in the control inventory at this cell's coordinate.
    /// </summary>
    public int? ControlId { get; init; }

    /// <summary>The cell-level control's kind (e.g. <c>Field</c>/<c>Label</c>), when <see cref="IsControlCell"/> is true.</summary>
    public string? ControlKind { get; init; }

    /// <summary>The cell-level control's <c>w:alias</c>, when <see cref="IsControlCell"/> is true.</summary>
    public string? Alias { get; init; }

    /// <summary>The cell-level control's binding XPath, when <see cref="IsControlCell"/> is true.</summary>
    public string? XPath { get; init; }

    /// <summary>The cell's visible text (whitespace-collapsed), excluding any nested-table content.</summary>
    public required string Text { get; init; }

    /// <summary>
    /// The <c>w:id</c>s of any content controls INSIDE this cell (excluding the cell-level wrapper itself
    /// and anything in a nested table) — e.g. a run-level field placed inside an ordinary cell. Ready to
    /// pass to <c>remove_control</c>.
    /// </summary>
    public required IReadOnlyList<int> InnerControlIds { get; init; }
}
