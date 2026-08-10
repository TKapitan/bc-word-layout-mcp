using System.Globalization;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Domain;

/// <summary>
/// Pins a tool-created table to the column widths it was asked for, by declaring the two
/// <c>w:tblPr</c> elements every corpus table carries: <c>w:tblW</c> (the table's total width) and
/// <c>w:tblLayout w:type="fixed"</c> (use the declared grid, do not recompute it).
/// </summary>
/// <remarks>
/// <para>
/// WHY IT MATTERS: without <c>w:tblLayout</c> the OOXML default is AUTOFIT, so Word is free to recompute
/// every column width from cell content, and a table's <c>w:tblGrid</c>/<c>w:tcW</c> values become a
/// preference rather than a specification. The <c>columnWidths</c> a caller passes to
/// <c>insert_table</c>/<c>insert_repeater_table</c> then hold only until Word opens the file: a 5-column line
/// table authored at <c>1500,4700,1200,1300,1300</c> read as near-uniform in Word (narrow columns' captions
/// wrapping rather than the columns keeping their proportions), and one plain Word save rewrote an
/// UNTOUCHED 2-column grid from <c>2800,3000</c> to <c>3177,3066</c> (GitHub issue #52). Dragging a column
/// divider in Word writes <c>tblLayout fixed</c> — Word's own repair for the table is exactly the element
/// that was missing.
/// </para>
/// <para>
/// OBSERVED, NOT INVENTED (ADR-0005): all four tables in <c>StandardPurchaseOrder.docx</c> carry both
/// elements. Three declare <c>w:tblW w:w="0" w:type="auto"</c> and one declares
/// <c>w:tblW w:w="10225" w:type="dxa"</c>; this helper emits the <c>dxa</c> form with the grid's own sum,
/// which is the shape that keeps the table self-describing (an <c>auto</c> total says nothing about the
/// width the caller actually chose) and the one the table-structure tools already maintain —
/// <c>TableStructureEditor.SyncTableWidthToGrid</c> recomputes a <c>dxa</c> total after every
/// <c>set_column_widths</c>/<c>insert_column</c>/<c>remove_column</c>, so the declared total cannot drift
/// out of step with the grid it describes.
/// </para>
/// <para>
/// CREATE-TIME ONLY. Nothing here retrofits an existing table: a captured layout's tables already declare
/// their own layout algorithm, and switching a stock table from autofit to fixed would change how BC renders
/// a document this tool was only asked to edit. The two callers are the two places that BUILD a table —
/// <see cref="LayoutEditor"/>'s plain-table insert and <see cref="SdtFactory"/>'s repeater table.
/// </para>
/// </remarks>
public static class FixedTableLayout
{
    /// <summary>
    /// Declares <paramref name="tblPr"/>'s total width as the sum of <paramref name="columnWidths"/> and its
    /// layout algorithm as fixed. Both are assigned through the SDK's typed properties rather than appended,
    /// so each lands at its schema-mandated position in <c>CT_TblPrBase</c>'s sequence (<c>w:tblW</c> before
    /// <c>w:tblBorders</c>, <c>w:tblLayout</c> after it) no matter what else the caller has already set.
    /// </summary>
    public static void ApplyTo(TableProperties tblPr, IReadOnlyList<int> columnWidths)
    {
        ArgumentNullException.ThrowIfNull(tblPr);
        ArgumentNullException.ThrowIfNull(columnWidths);

        var total = 0;
        foreach (var width in columnWidths)
        {
            total += width;
        }

        tblPr.TableWidth = new TableWidth
        {
            Type = TableWidthUnitValues.Dxa,
            Width = total.ToString(CultureInfo.InvariantCulture),
        };
        tblPr.TableLayout = new TableLayout { Type = TableLayoutValues.Fixed };
    }
}
