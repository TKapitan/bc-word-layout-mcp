using System.Globalization;
using BcWordLayout.Domain.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Domain;

/// <summary>
/// Describes every table in a layout structurally — grid dimensions, rows, and cells, together with which
/// content control (if any) owns each cell/row — and, in the same pass, records the (table, row, column)
/// coordinates of EVERY content control that lives inside a table so <see cref="LayoutReader"/> can attach
/// those coordinates to its flat control inventory.
/// </summary>
/// <remarks>
/// <para>
/// Table/row/cell enumeration routes entirely through <see cref="TableGridNavigator"/> — the single source
/// of truth also used by <see cref="LocationResolver"/>'s <c>tableCell</c> addressing and by table-structure
/// editing — so tables are numbered 0-based, per part, in the SAME document order everywhere (nested tables
/// flattened into that numbering, each getting its own index), and a table reported here can be targeted by
/// the insert tools' <c>tableIndex</c> directly, by construction rather than by hand-kept parity.
/// </para>
/// <para>
/// Each table's own rows/cells are read from its DIRECT children only — a control living inside a NESTED
/// table is attributed to that nested table's own entry (its innermost enclosing <c>w:tbl</c>), never to the
/// outer cell that happens to contain the nested table, so every control is assigned to exactly one cell.
/// </para>
/// </remarks>
public static class TableStructureReader
{
    /// <summary>The (table, row, column) coordinates of a control within a part. <c>Col</c> is null for a row-level control.</summary>
    public readonly record struct Coord(int TableIndex, int RowIndex, int? ColIndex);

    /// <summary>
    /// Reads the table structures of every part plus a reference-keyed map from each in-table
    /// <see cref="SdtElement"/> to its <see cref="Coord"/>. The map keys are the live element instances in
    /// <paramref name="doc"/>, so callers must look controls up while that same tree is still open.
    /// </summary>
    public static (IReadOnlyList<TableStructure> Tables, IReadOnlyDictionary<SdtElement, Coord> Coords) Read(WordprocessingDocument doc)
    {
        ArgumentNullException.ThrowIfNull(doc);

        var main = doc.MainDocumentPart
            ?? throw new InvalidDataException("Layout has no main document part.");

        var tables = new List<TableStructure>();
        var coords = new Dictionary<SdtElement, Coord>(ReferenceEqualityComparer.Instance);

        foreach (var (root, partName) in PartWalker.ContentParts(main))
        {
            var partTables = TableGridNavigator.Tables(root);
            for (var t = 0; t < partTables.Count; t++)
            {
                tables.Add(ReadTable(partTables[t], partName, t, coords));
            }
        }

        return (tables, coords);
    }

    private static TableStructure ReadTable(Table table, string partName, int tableIndex, Dictionary<SdtElement, Coord> coords)
    {
        // Declared widths (for GridColumnWidths, which is documented as "empty when none are present" and
        // is NOT meant to be index-aligned with the grid) stay filtered to width-bearing gridCols. But
        // ColumnCount must count EVERY w:gridCol, width-bearing or not: a w:gridCol without w:w is
        // schema-legal, and every editor (set_column_widths/remove_column/insert_column via
        // TableGridNavigator.GridColumnCount) already treats the grid that way. Reporting fewer here would
        // silently break get_layout_info's documented "pass columnCount from here" contract.
        var widths = table.GetFirstChild<TableGrid>()?.Elements<GridColumn>()
            .Select(g => int.TryParse(g.Width?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var w) ? (int?)w : null)
            .Where(w => w.HasValue)
            .Select(w => w!.Value)
            .ToList() ?? new List<int>();
        var columnCount = TableGridNavigator.GridColumnCount(table);

        // Rows in document order (a plain w:tr, OR a row-level sdt sitting where a w:tr would) — the same
        // enumeration TableGridNavigator.Rows exposes to every other row/cell walker.
        var rowSlots = TableGridNavigator.Rows(table);
        var rows = new List<TableRowInfo>(rowSlots.Count);

        for (var r = 0; r < rowSlots.Count; r++)
        {
            rows.Add(ReadRow(table, rowSlots[r], tableIndex, r, coords));
        }

        return new TableStructure
        {
            Part = partName,
            TableIndex = tableIndex,
            RowCount = rows.Count,
            ColumnCount = columnCount,
            GridColumnWidths = widths,
            Rows = rows,
        };
    }

    private static TableRowInfo ReadRow(
        Table table, TableGridNavigator.RowSlot rowSlot, int tableIndex, int rowIndex, Dictionary<SdtElement, Coord> coords)
    {
        // A row-level control owns the whole row: record EVERY wrapper in its chain (e.g. a BC repeater's
        // repeatingSection wrapping a repeatingSectionItem) at (table, row, null) — not just the outermost
        // one — then read the cells from the innermost w:tr the chain ultimately holds. The OUTERMOST
        // wrapper's id is reported as the row's own ControlId.
        if (rowSlot.IsControlRow)
        {
            var rowCoord = new Coord(tableIndex, rowIndex, null);
            foreach (var wrapper in rowSlot.Wrappers)
            {
                coords[wrapper] = rowCoord;
            }
        }

        var cells = new List<TableCellInfo>();
        if (rowSlot.InnerRow is not null)
        {
            var cellSlots = TableGridNavigator.Cells(rowSlot.InnerRow);
            for (var c = 0; c < cellSlots.Count; c++)
            {
                cells.Add(ReadCell(table, cellSlots[c], tableIndex, rowIndex, c, coords));
            }
        }

        return new TableRowInfo
        {
            RowIndex = rowIndex,
            IsControlRow = rowSlot.IsControlRow,
            ControlId = rowSlot.ControlId,
            Cells = cells,
        };
    }

    private static TableCellInfo ReadCell(
        Table table, TableGridNavigator.CellSlot cellSlot, int tableIndex, int rowIndex, int colIndex, Dictionary<SdtElement, Coord> coords)
    {
        var coord = new Coord(tableIndex, rowIndex, colIndex);

        TableCell? cell = cellSlot.InnerCell;
        bool isControlCell;
        int? cellControlId = null;
        string? controlKind = null, alias = null, xpath = null;

        if (cellSlot.CellChild is SdtCell sdtCell)
        {
            isControlCell = true;
            cellControlId = SdtInspector.ReadControlId(sdtCell);
            controlKind = SdtInspector.ClassifyControlKind(sdtCell).ToString();
            alias = SdtInspector.ReadAlias(sdtCell);
            xpath = SdtInspector.ReadXPath(sdtCell);
            coords[sdtCell] = coord;
        }
        else
        {
            isControlCell = false;
        }

        // Controls and text that belong to THIS cell (not to a nested table inside it): filter on the
        // innermost enclosing table so nested-table content is attributed to its own table entry instead.
        var innerIds = new List<int>();
        var text = string.Empty;
        if (cell is not null)
        {
            foreach (var sdt in cell.Descendants<SdtElement>())
            {
                if (ReferenceEquals(NearestTable(sdt), table))
                {
                    coords.TryAdd(sdt, coord);
                    var id = SdtInspector.ReadControlId(sdt);
                    if (id.HasValue && (!isControlCell || !ReferenceEquals(sdt, cellSlot.CellChild)))
                    {
                        innerIds.Add(id.Value);
                    }
                }
            }

            text = CollapseWhitespace(string.Concat(
                cell.Descendants<Text>()
                    .Where(t => ReferenceEquals(NearestTable(t), table))
                    .Select(t => t.Text)));
        }

        return new TableCellInfo
        {
            ColIndex = colIndex,
            IsControlCell = isControlCell,
            ControlId = cellControlId,
            ControlKind = controlKind,
            Alias = alias,
            XPath = xpath,
            Text = text,
            InnerControlIds = innerIds,
        };
    }

    private static Table? NearestTable(OpenXmlElement el) => el.Ancestors<Table>().FirstOrDefault();

    private static string CollapseWhitespace(string s) =>
        string.Join(' ', s.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
}
