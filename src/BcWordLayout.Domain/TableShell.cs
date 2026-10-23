using BcWordLayout.Domain.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Domain;

/// <summary>
/// The table-level shape both new-table tools emit — <c>insert_table</c>'s plain table
/// (<see cref="LayoutEditor.InsertPlainTable"/>) and <c>insert_repeater_table</c>'s line table
/// (<see cref="SdtFactory.BuildRepeaterTable"/>) — so their shared <c>borders</c> and <c>tableStyle</c> options
/// mean exactly the same thing on both (GitHub issue #131): the <c>w:tblPr</c>, the <c>headerRule</c> rule under
/// the first row, and the check that a style reference names a table style the layout defines.
/// </summary>
internal static class TableShell
{
    /// <summary>
    /// The new table's <c>w:tblPr</c>: a <c>w:tblStyle</c> reference when <paramref name="tableStyle"/> is given,
    /// the grid pinned to <paramref name="widths"/> (see <see cref="FixedTableLayout"/>), and — for
    /// <see cref="TableBorderLook.Grid"/> only — an explicit single-line border on every edge and between every
    /// cell. <see cref="TableBorderLook.None"/> and <see cref="TableBorderLook.HeaderRule"/> emit no
    /// <c>w:tblBorders</c> at all, exactly like every corpus table: the header rule is drawn per cell
    /// (<see cref="DrawHeaderRule"/>). Assigned through the SDK's typed properties throughout, so each element
    /// lands at its schema position whatever order they are set in.
    /// </summary>
    internal static TableProperties Properties(IReadOnlyList<int> widths, TableBorderLook borders, string? tableStyle)
    {
        var tblPr = new TableProperties();
        if (!string.IsNullOrEmpty(tableStyle))
        {
            tblPr.TableStyle = new TableStyle { Val = tableStyle };
        }

        // Pin the grid the caller asked for; without this Word autofits and recomputes every width from cell
        // content (see FixedTableLayout).
        FixedTableLayout.ApplyTo(tblPr, widths);

        if (borders == TableBorderLook.Grid)
        {
            tblPr.TableBorders = new TableBorders
            {
                TopBorder = new TopBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
                LeftBorder = new LeftBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
                BottomBorder = new BottomBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
                RightBorder = new RightBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
                InsideHorizontalBorder =
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
                InsideVerticalBorder =
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
            };
        }

        return tblPr;
    }

    /// <summary>
    /// <see cref="TableBorderLook.HeaderRule"/>'s one rule: a ½-pt bottom edge on each of <paramref name="cells"/>
    /// (the first row's), drawn through <see cref="TableStructureEditor.ApplyCellBorders"/> so it is
    /// byte-identical to the rule <c>set_cell_format</c> draws for a bottom edge.
    /// </summary>
    internal static void DrawHeaderRule(IEnumerable<TableCell> cells)
    {
        foreach (var cell in cells)
        {
            TableStructureEditor.ApplyCellBorders(cell, new CellBorderOptions { Bottom = true });
        }
    }

    /// <summary>
    /// Returns <paramref name="styleId"/> when it is the id of a TABLE style the layout's styles part defines,
    /// else refuses the call, listing the table styles the layout does define. A dangling <c>w:tblStyle</c> is
    /// valid OOXML that Word and Business Central silently ignore, so writing one would leave the caller
    /// believing the table is styled when every renderer draws its own defaults; <c>validate_layout</c>'s
    /// <c>table-style-resolves</c> warning still covers the references no tool wrote. The same rule
    /// <c>format.characterStyle</c> follows (<see cref="RunFormatter"/>). Style ids are case-sensitive.
    /// </summary>
    internal static string CheckTableStyle(WordprocessingDocument doc, string styleId, string paramName)
    {
        var styles = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles?.Elements<Style>().ToList() ?? [];
        var match = styles.FirstOrDefault(s => string.Equals(s.StyleId?.Value, styleId, StringComparison.Ordinal));
        if (match?.Type?.Value == StyleValues.Table)
        {
            return styleId;
        }

        var tableStyles = styles
            .Where(s => s.Type?.Value == StyleValues.Table && s.StyleId?.Value is { Length: > 0 })
            .Select(s => s.StyleId!.Value!)
            .Where(id => id != "TableNormal")
            .ToList();
        var available = tableStyles.Count == 0
            ? "This layout defines no table styles: start from a templatePath that defines the style, or draw the "
              + "rules with borders and set_cell_format instead."
            : $"Table styles this layout defines: {string.Join(", ", tableStyles.Select(s => $"'{s}'"))}.";
        var reason = match is null
            ? "the layout's styles part does not define it"
            : $"it is a {match.Type?.InnerText ?? "typeless"} style, not a table style";

        throw new ArgumentException(
            $"tableStyle '{styleId}' cannot be applied: {reason}, so the reference would do nothing. " + available,
            paramName);
    }
}
