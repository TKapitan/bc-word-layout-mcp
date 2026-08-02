namespace BcWordLayout.Domain.Models;

/// <summary>
/// Describes a single mutation performed by <see cref="BcWordLayout.Domain.LayoutEditor"/> — enough detail
/// for a caller (typically an MCP editing tool) to report exactly what changed without re-reading the
/// document. Every field beyond <see cref="Operation"/>/<see cref="ControlId"/>/<see cref="Part"/>/
/// <see cref="Summary"/> is best-effort (null when not known/applicable — e.g. an unbound control has no
/// <see cref="XPath"/>).
/// </summary>
public sealed class EditResult
{
    /// <summary>Stable operation name, e.g. <c>insert_field</c>, <c>insert_label</c>, <c>remove_control</c>.</summary>
    public required string Operation { get; init; }

    /// <summary>The affected control's <c>w:id</c>.</summary>
    public required int ControlId { get; init; }

    /// <summary>The control's <c>w:alias</c> value, when known (e.g. <c>#Nav: /Header/CustomerAddress1</c>).</summary>
    public string? Alias { get; init; }

    /// <summary>The control's binding XPath, when known/applicable.</summary>
    public string? XPath { get; init; }

    /// <summary>
    /// The control's classification (a <see cref="ControlKind"/> name, e.g. <c>Field</c>/<c>Label</c>),
    /// when known.
    /// </summary>
    public string? Kind { get; init; }

    /// <summary>
    /// Number of data columns in an inserted repeater table (<c>insert_repeater_table</c> only); null for
    /// every other operation.
    /// </summary>
    public int? ColumnCount { get; init; }

    /// <summary>
    /// The 0-based, per-part, document-order index of a table an operation created
    /// (<c>insert_repeater_table</c> only; null for every other operation) — the SAME coordinate
    /// <c>get_layout_info</c>'s <c>tables[]</c> and every <c>tableCell</c> location use, so a follow-up edit
    /// can address the new table without re-reading the layout.
    /// </summary>
    public int? TableIndex { get; init; }

    /// <summary>
    /// The 0-based row index, within <see cref="TableIndex"/>, of an inserted repeater table's DATA row (the
    /// row wrapped by the repeating-section control); null for every other operation. Paired with
    /// <see cref="TableIndex"/> this is exactly the coordinate a NESTED repeater — or any per-row follow-up
    /// edit — must target, which is otherwise only discoverable by re-reading the layout and reasoning about
    /// which row is the template.
    /// </summary>
    public int? DataRowIndex { get; init; }

    /// <summary>Which OOXML part the control lives (or lived) in, e.g. <c>document.xml</c>, <c>header1.xml</c>.</summary>
    public required string Part { get; init; }

    /// <summary>Human-readable one-line summary of the change.</summary>
    public required string Summary { get; init; }
}

/// <summary>
/// Optional formatting for <see cref="BcWordLayout.Domain.CellTextEditor.SetCellText"/>: every null member
/// means "keep whatever the cell already had" (the pre-format behavior exactly), so a plain re-label never
/// disturbs existing styling. Needed because a cell in a freshly authored plain table
/// (<c>insert_table</c>) has NO styling to inherit — caption cells need bold, amount columns need right
/// alignment, a document title needs a bigger size.
/// </summary>
public sealed record CellTextFormat
{
    /// <summary>True adds bold, false removes it, null keeps the inherited weight.</summary>
    public bool? Bold { get; init; }

    /// <summary><c>left</c>/<c>center</c>/<c>right</c> (case-insensitive); null keeps the inherited alignment.</summary>
    public string? Alignment { get; init; }

    /// <summary>Font size in points (4–96, halves allowed); null keeps the inherited size.</summary>
    public double? FontSizePoints { get; init; }
}

/// <summary>
/// Describes a single plain-text table-cell mutation performed by
/// <see cref="BcWordLayout.Domain.CellTextEditor"/> (<c>set_cell_text</c>/<c>clear_cell_text</c>). Unlike
/// <see cref="EditResult"/>, a cell-text edit targets a physical cell by its (table, row, column)
/// coordinates rather than a control's <c>w:id</c> — a plain-text cell has no control to identify.
/// </summary>
public sealed class CellEditResult
{
    /// <summary>Stable operation name, <c>set_cell_text</c> or <c>clear_cell_text</c>.</summary>
    public required string Operation { get; init; }

    /// <summary>Which OOXML part the cell lives in, e.g. <c>document.xml</c>, <c>header1.xml</c>.</summary>
    public required string Part { get; init; }

    /// <summary>0-based table index (per part, document order) — the same index <c>get_layout_info</c> reports.</summary>
    public required int TableIndex { get; init; }

    /// <summary>0-based row index within the table.</summary>
    public required int Row { get; init; }

    /// <summary>0-based cell/column index within the row.</summary>
    public required int Col { get; init; }

    /// <summary>The cell's full visible text BEFORE the edit (all runs concatenated).</summary>
    public required string PreviousText { get; init; }

    /// <summary>The cell's visible text AFTER the edit — the new text for <c>set_cell_text</c>, empty for <c>clear_cell_text</c>.</summary>
    public required string NewText { get; init; }

    /// <summary>Human-readable one-line summary of the change.</summary>
    public required string Summary { get; init; }
}

/// <summary>
/// Describes a single table-STRUCTURE mutation performed by
/// <see cref="BcWordLayout.Domain.TableStructureEditor"/> (<c>set_column_widths</c>, <c>insert_column</c>,
/// <c>remove_column</c>, …). Unlike <see cref="EditResult"/> (a control identified by <c>w:id</c>) or
/// <see cref="CellEditResult"/> (a single cell by coordinates), a structure edit acts on a whole table
/// column, so it reports the affected column index and the grid column count before/after.
/// </summary>
public sealed class TableEditResult
{
    /// <summary>Stable operation name, e.g. <c>set_column_widths</c>, <c>insert_column</c>, <c>remove_column</c>.</summary>
    public required string Operation { get; init; }

    /// <summary>Which OOXML part the table lives in, e.g. <c>document.xml</c>, <c>header1.xml</c>.</summary>
    public required string Part { get; init; }

    /// <summary>0-based table index (per part, document order) — the same index <c>get_layout_info</c> reports.</summary>
    public required int TableIndex { get; init; }

    /// <summary>The 0-based GRID column index the operation added/removed/targeted; null for a whole-table op like a width set.</summary>
    public int? ColumnIndex { get; init; }

    /// <summary>Number of table rows the edit actually touched (cells added/removed/resized).</summary>
    public required int RowsAffected { get; init; }

    /// <summary>Grid column count before the edit.</summary>
    public required int ColumnCountBefore { get; init; }

    /// <summary>Grid column count after the edit.</summary>
    public required int ColumnCountAfter { get; init; }

    /// <summary>Human-readable one-line summary of the change.</summary>
    public required string Summary { get; init; }
}

/// <summary>
/// One cell of a nested detail row (<see cref="BcWordLayout.Domain.LayoutEditor.InsertRepeaterRow"/>):
/// how many grid columns of the PARENT table it spans, which of the child data item's leaf columns it
/// carries (label-shaped names become label controls, chained inline in one paragraph — the
/// corpus-verified shape; empty for a spacer cell), and an optional paragraph alignment.
/// </summary>
public sealed record RepeaterRowCell
{
    /// <summary>Grid columns of the parent table this cell spans (≥1); all cells must sum to the parent's grid count.</summary>
    public int Span { get; init; } = 1;

    /// <summary>Leaf column names of the child data item rendered inline in this cell, in order; empty = spacer cell.</summary>
    public IReadOnlyList<string> Columns { get; init; } = [];

    /// <summary><c>left</c>/<c>center</c>/<c>right</c>; null keeps the default.</summary>
    public string? Alignment { get; init; }
}

/// <summary>The kind of column <see cref="BcWordLayout.Domain.TableStructureEditor.InsertColumn"/> adds.</summary>
public enum InsertColumnMode
{
    /// <summary>A bound FIELD control cell (built exactly as <c>insert_field</c> does).</summary>
    Field,

    /// <summary>A bound LABEL control cell (built exactly as <c>insert_label</c> does).</summary>
    Label,

    /// <summary>A plain-text cell (no binding): an optional header text plus empty data cells.</summary>
    PlainText,
}

/// <summary>Options controlling <see cref="BcWordLayout.Domain.TableStructureEditor.InsertColumn"/>.</summary>
public sealed class InsertColumnOptions
{
    /// <summary>What kind of new column to add.</summary>
    public required InsertColumnMode Mode { get; init; }

    /// <summary>
    /// For <see cref="InsertColumnMode.Field"/>/<see cref="InsertColumnMode.Label"/>: the full dataset path
    /// to bind the data cell to (e.g. <c>/Header/Line/Discount_Line</c>), exactly as <c>insert_field</c>/
    /// <c>insert_label</c> take. Ignored for <see cref="InsertColumnMode.PlainText"/>.
    /// </summary>
    public string? DataPath { get; init; }

    /// <summary>Optional: bind the header cell to this label-column dataset path (by default a <c>*Lbl</c>/<c>*_Lbl</c> path — see <c>BcWordLayout.Domain.LabelConvention</c>).</summary>
    public string? HeaderLabelPath { get; init; }

    /// <summary>Optional: static header-cell text. For <see cref="InsertColumnMode.PlainText"/> this is the header.</summary>
    public string? HeaderText { get; init; }

    /// <summary>Optional explicit new-column width in twips; when omitted, the mean of the existing grid columns.</summary>
    public int? Width { get; init; }
}

/// <summary>
/// The border treatment a freshly authored repeater table gets. Real BC line-items tables are NOT drawn as
/// a grid: they carry no table-level <c>w:tblBorders</c> at all, and their look comes entirely from per-cell
/// <c>w:tcBorders</c> — a single rule under the header row, and (where a totals block follows) a rule above
/// it. Corpus-verified: every <c>w:tcBorders</c> across the six captured layouts is one of exactly two
/// shapes, <c>&lt;w:bottom w:val="single" w:sz="4" w:color="auto"/&gt;</c> or the same with <c>w:top</c>.
/// </summary>
public enum TableBorderLook
{
    /// <summary>
    /// The BC-native look (default): no table-level border grid, one ½-pt rule under the header row. The
    /// remaining rules a real BC document has (above a totals block, say) are per-cell decisions the author
    /// makes afterwards with <c>set_cell_borders</c>.
    /// </summary>
    Bc,

    /// <summary>
    /// An explicit single-line border on every table edge and between every cell — visible without depending
    /// on an existing style, and the right choice for a genuinely grid-shaped table, but not what a BC
    /// line-items table looks like.
    /// </summary>
    Grid,
}

/// <summary>
/// Which edges of a table cell <see cref="BcWordLayout.Domain.TableStructureEditor.SetCellBorders"/> should
/// draw (or clear), and how thick. This is the per-cell half of the BC table look (see
/// <see cref="TableBorderLook"/>): the rule under a header row, the rule above a totals block, an underline on one
/// summary cell. <see cref="Remove"/> flips the meaning of the selected edges from "draw" to "clear".
/// </summary>
public sealed record CellBorderOptions
{
    /// <summary>The BC-standard rule thickness in eighths of a point (4 = ½ pt) — the only value the corpus uses.</summary>
    public const int DefaultSizeEighthPoints = 4;

    /// <summary>Draw (or, with <see cref="Remove"/>, clear) the cell's TOP edge.</summary>
    public bool Top { get; init; }

    /// <summary>Draw (or, with <see cref="Remove"/>, clear) the cell's BOTTOM edge.</summary>
    public bool Bottom { get; init; }

    /// <summary>Draw (or, with <see cref="Remove"/>, clear) the cell's LEFT edge.</summary>
    public bool Left { get; init; }

    /// <summary>Draw (or, with <see cref="Remove"/>, clear) the cell's RIGHT edge.</summary>
    public bool Right { get; init; }

    /// <summary>
    /// When true the selected edges are explicitly cleared (<c>w:val="nil"</c>) rather than drawn — an
    /// explicit nil, not a deleted element, so the edge also stays hidden in a table that DOES declare a
    /// table-level border grid for its cells to inherit from.
    /// </summary>
    public bool Remove { get; init; }

    /// <summary>Rule thickness in eighths of a point (2–96); see <see cref="DefaultSizeEighthPoints"/>.</summary>
    public int SizeEighthPoints { get; init; } = DefaultSizeEighthPoints;

    /// <summary>True when at least one edge is selected (a call selecting none changes nothing and is rejected).</summary>
    public bool AnyEdge => Top || Bottom || Left || Right;
}

/// <summary>
/// Options controlling <see cref="BcWordLayout.Domain.SdtFactory.BuildRepeaterTable"/> /
/// <see cref="BcWordLayout.Domain.LayoutEditor.InsertRepeaterTable"/>.
/// </summary>
public sealed class RepeaterTableOptions
{
    /// <summary>
    /// The table's border treatment — <see cref="TableBorderLook.Bc"/> (the default, matching real BC line-items
    /// tables) or <see cref="TableBorderLook.Grid"/>. See <see cref="TableBorderLook"/> for the corpus evidence.
    /// </summary>
    public TableBorderLook Look { get; init; } = TableBorderLook.Bc;

    /// <summary>
    /// When true (the default), each header cell binds to its column's label column (by default suffixed
    /// <c>*Lbl</c>/<c>*_Lbl</c> — see <c>BcWordLayout.Domain.LabelConvention</c>) when one can be found (see
    /// <see cref="BcWordLayout.Domain.SdtFactory.BuildRepeaterTable"/>'s remarks for the exact lookup
    /// order); otherwise (or when no label column is found) the header cell is static humanized text
    /// instead. When false, every header cell is static humanized text.
    /// </summary>
    public bool HeaderFromLabels { get; init; } = true;

    /// <summary>Optional Word table style name to reference via <c>w:tblStyle</c> (e.g. <c>"TableGrid"</c>).</summary>
    public string? TableStyle { get; init; }

    /// <summary>
    /// Optional explicit column widths in twips (1/20 pt), one per column, applied to both <c>w:tblGrid</c>
    /// and each column's cell width. Must have exactly one entry per column when supplied. When omitted,
    /// every column gets the same even default width.
    /// </summary>
    public IReadOnlyList<int>? ColumnWidths { get; init; }

    /// <summary>
    /// Optional per-column paragraph alignments (<c>left</c>/<c>center</c>/<c>right</c>, case-insensitive),
    /// one per column, applied to BOTH the header cell and the data cell of each column — real BC line
    /// tables right-align their quantity/price/amount columns. Must have exactly one entry per column when
    /// supplied; when omitted, cells keep the default (left).
    /// </summary>
    public IReadOnlyList<string>? ColumnAlignments { get; init; }
}
