using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Domain;

/// <summary>
/// Which structural rule a <see cref="TableGridViolation"/> breaks — the STABLE part of a violation's
/// identity (see <see cref="TableGridViolation"/>'s remarks for why this, and not the free-text
/// <see cref="TableGridViolation.Reason"/>, is what <c>GuardMutate</c>'s BEFORE/AFTER diff keys on).
/// </summary>
public enum TableGridViolationKind
{
    /// <summary>A row's total cell coverage (Σ<c>gridSpan</c> plus any <c>gridBefore</c>/<c>gridAfter</c>) does not equal the table's declared <c>w:tblGrid</c> column count.</summary>
    CoverageMismatch,

    /// <summary>A cell holds no block-level content (its <c>w:tc</c> has nothing but <c>w:tcPr</c>, or its cell-level control wraps no <c>w:tc</c> at all).</summary>
    EmptyCell,
}

/// <summary>
/// One table row whose cell layout is inconsistent with its <c>w:tblGrid</c> — either its cells cover a
/// different number of grid columns than the grid declares, or one of its cells is empty of block content.
/// Both are Word-visible table corruptions that <see cref="DocumentFormat.OpenXml.Validation.OpenXmlValidator"/>
/// accepts silently (a <c>w:tr</c> may schema-legally hold any number of <c>w:tc</c>, and a
/// <c>w:tc</c>'s content model does not force a paragraph), which is exactly why this dedicated guard
/// exists alongside the structural gate — the same relationship <see cref="PlainTextNestingGuard"/> has.
/// </summary>
/// <remarks>
/// <para>
/// IDENTITY vs. DISPLAY. <c>GuardMutate</c> (in
/// <c>BcWordLayout.McpHost.Tools.ToolGuards</c>) diffs a BEFORE-edit set of these against an AFTER-edit set
/// to decide whether an edit introduced NEW damage — a pre-existing violation must stay editable, never
/// misreported as caused by the edit in progress (the same promise the OpenXmlValidator and
/// <see cref="PlainTextNestingGuard"/> baseline-diffs make). Earlier, that diff keyed on <see cref="Describe"/>'s
/// full string, which embeds the LIVE <c>gridCount</c>/coverage numbers — so <c>insert_column</c>/
/// <c>remove_column</c>, which change every row's expected grid count by design, made a pre-existing
/// violation's description reword itself ("declares 9" → "declares 10") and read as newly introduced.
/// </para>
/// <para>
/// The diff (see <c>ToolGuards.DiffTableGridViolations</c>) now keys on <c>(</c><see cref="Part"/><c>,</c>
/// <see cref="TableIndex"/><c>,</c> <see cref="Kind"/><c>)</c> — deliberately NOT <see cref="RowIndex"/> and
/// NOT the numbers embedded in <see cref="Reason"/> — and REJECTS ONLY WHEN THE COUNT of violations sharing a
/// key goes up, rather than requiring exact set membership. This was chosen over the more literal
/// "<c>(Part, TableIndex, RowIndex, Kind)</c>, exact set" reading because a row can legitimately shift its
/// <see cref="RowIndex"/> without being touched at all: <c>remove_control</c> on a row-level control (e.g. a
/// repeater's own row template) deletes that whole physical row, shifting every later row's index down by
/// one. Under an exact-membership diff, a pre-existing violation on a row below the deleted one would appear
/// to vanish at its old index and reappear "new" at its shifted index — falsely rejecting an edit that never
/// touched that row. Dropping <see cref="RowIndex"/> from the key and comparing COUNTS per
/// <c>(Part, TableIndex, Kind)</c> absorbs that shift for free: the same table still reports the same NUMBER
/// of violations of that kind, just at a different row, so no increase is observed. <see cref="Reason"/>/
/// <see cref="Describe"/> remain purely for the human-readable error message; they carry no identity weight
/// any more.
/// </para>
/// <para>
/// ACCEPTED IMPRECISION: <see cref="TableIndex"/> itself can shift too — <c>insert_repeater_table</c> adds a
/// whole new <c>w:tbl</c>, and if it lands BEFORE an already-damaged table in the same part, every table
/// after it (including the damaged one) shifts up by one index, which this key does not tolerate (it would
/// misfire exactly like the RowIndex case, just one level up). This residual gap is accepted rather than
/// solved: no current mutating tool ever REMOVES a whole table (only adds one), so it can only bite in the
/// narrow case of inserting a new repeater table at a location that precedes a pre-existing damaged table in
/// the same part — considerably rarer than the reported defect (every <c>insert_column</c>/<c>remove_column</c>
/// call on ANY damaged table), and easy to avoid by inserting after the point of damage. Revisit if a future
/// tool removes tables, or if this proves to bite in practice.
/// </para>
/// <para>
/// ACCEPTED IMPRECISION #2 (the count rule's false-NEGATIVE class): comparing counts per
/// <c>(Part, TableIndex, Kind)</c> is strictly weaker than the old exact-set diff in one way — an edit that
/// FIXES one violation of a kind while INTRODUCING a different one of the SAME kind in the SAME table leaves
/// the count unchanged and is accepted, where the exact-set diff would have rejected it. No current mutating
/// operation can produce such a swap: <c>insert_column</c> preserves every row's coverage-vs-grid delta
/// exactly, <c>remove_column</c> skips already-inconsistent rows and can only reduce the mismatch count,
/// <c>set_column_widths</c> is width-only, and <c>merge_cells</c>/<c>split_cells</c> preserve a single row's
/// coverage — so the swap is only reachable through a FUTURE tool bug, which this backstop would then miss.
/// Accepted because the guard is the LAST line behind OpenXmlValidator and each tool's own
/// correct-by-construction editing logic, and the swap-blindness is the direct price of absorbing the
/// legitimate index/count shifts above. Revisit if a future table operation can both fix and break rows in
/// one edit.
/// </para>
/// </remarks>
public readonly record struct TableGridViolation(string Part, int TableIndex, int RowIndex, TableGridViolationKind Kind, string Reason)
{
    /// <summary>
    /// A human-readable one-line description for error messages ONLY — see the type's own remarks for why
    /// this is no longer what <c>GuardMutate</c>'s BEFORE/AFTER diff keys on (it embeds live counts that
    /// legitimately change under column edits).
    /// </summary>
    public string Describe() => $"table {TableIndex} row {RowIndex} in {Part}: {Reason}";
}

/// <summary>
/// Detects table rows whose cell layout is inconsistent with the table's <c>w:tblGrid</c>. For every row
/// (resolving repeater/cell wrappers via <see cref="TableGridNavigator"/>), the sum of its cells'
/// <c>w:gridSpan</c> values plus any <c>w:gridBefore</c>/<c>w:gridAfter</c> must equal the grid's column
/// count; and every cell must hold at least one block-level child.
/// </summary>
/// <remarks>
/// WHY THIS EXISTS: the structural table-editing tools (<c>set_column_widths</c>, <c>insert_column</c>,
/// <c>remove_column</c>, …) add/remove/resize cells and grid columns. A bug that left a row covering fewer
/// (or more) grid columns than the grid declares, or that emptied a <c>w:tc</c>, would produce a document
/// Word renders as a broken/corrupt table — yet it is well-formed OOXML, so the pre-save
/// <see cref="DocumentFormat.OpenXml.Validation.OpenXmlValidator"/> gate in <c>GuardMutate</c> never sees
/// it (the same blind spot <see cref="PlainTextNestingGuard"/> and <c>LayoutEditor.EnsureCellNotEmpty</c>
/// already work around). Diffed BEFORE-vs-AFTER in <c>GuardMutate</c>, this guard rejects any edit that
/// desyncs a table's cells from its grid before it can reach disk — so a structural table edit is
/// mechanically prevented from corrupting the grid rather than relied upon to be correct by construction.
/// </remarks>
public static class TableGridConsistencyGuard
{
    /// <summary>
    /// Returns every grid-inconsistent row anywhere in <paramref name="doc"/> (main body, then each header,
    /// then each footer — the same part order the other guards/scans use). Empty when every table is
    /// consistent. Tables with no <c>w:tblGrid</c> are skipped (there is no declared column count to
    /// check against).
    /// </summary>
    public static IReadOnlyList<TableGridViolation> Find(WordprocessingDocument doc)
    {
        ArgumentNullException.ThrowIfNull(doc);

        var main = doc.MainDocumentPart
            ?? throw new InvalidDataException("Layout has no main document part.");

        var found = new List<TableGridViolation>();

        foreach (var (root, part) in PartWalker.ContentParts(main))
        {
            var tables = root.Descendants<Table>().ToList();
            for (var t = 0; t < tables.Count; t++)
            {
                CheckTable(tables[t], part, t, found);
            }
        }

        return found;
    }

    private static void CheckTable(Table table, string part, int tableIndex, List<TableGridViolation> found)
    {
        var gridCount = TableGridNavigator.GridColumnCount(table);
        if (gridCount == 0)
        {
            return; // No declared grid to check against.
        }

        var rows = TableGridNavigator.Rows(table);
        for (var r = 0; r < rows.Count; r++)
        {
            var innerRow = rows[r].InnerRow;
            if (innerRow is null)
            {
                continue; // A row-level control wrapping no w:tr has no cells to check.
            }

            var cells = TableGridNavigator.Cells(innerRow);
            var (before, after) = TableGridNavigator.GridEdges(innerRow);
            var coverage = before + after + cells.Sum(c => c.GridSpan);
            if (coverage != gridCount)
            {
                found.Add(new TableGridViolation(
                    part, tableIndex, r, TableGridViolationKind.CoverageMismatch,
                    $"cells cover {coverage} grid column(s) (ΣgridSpan {cells.Sum(c => c.GridSpan)}"
                    + $"{(before + after > 0 ? $" + gridBefore/After {before + after}" : string.Empty)}) "
                    + $"but the grid declares {gridCount}"));
            }

            for (var c = 0; c < cells.Count; c++)
            {
                var inner = cells[c].InnerCell;
                if (inner is not null && !inner.ChildElements.Any(e => e is not TableCellProperties))
                {
                    found.Add(new TableGridViolation(
                        part, tableIndex, r, TableGridViolationKind.EmptyCell,
                        $"cell {c} is empty (has no block-level content)"));
                }
            }
        }
    }
}
