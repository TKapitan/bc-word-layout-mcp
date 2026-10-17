using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

/// <summary>
/// A deliberately NAIVE second walker over a table's rows and cells, for assertions that must not share
/// <c>TableGridNavigator</c>'s blind spots. Every check built on the navigator — the tools, the
/// <c>TableGridConsistencyGuard</c>, and tests that ask either of them — inherits whatever the navigator
/// cannot see, which is how a column edit could leave the nested rows of a repeater item on the old grid
/// while every test passed (GitHub issue #80), and how a cell wrapped in two cell-level controls stayed
/// invisible (issue #85).
/// </summary>
/// <remarks>
/// The definitions are the plain OOXML ones and nothing more: a table's rows are EVERY <c>w:tr</c> whose
/// nearest enclosing <c>w:tbl</c> is that table, however deep the row-level controls around it go; a row's
/// cells are every <c>w:tc</c> whose nearest enclosing <c>w:tr</c> is that row, however many cell-level
/// controls wrap it. No wrapper is resolved, classified or skipped.
/// </remarks>
internal static class IndependentGridWalk
{
    /// <summary>Every <c>w:tr</c> whose nearest enclosing <c>w:tbl</c> is <paramref name="table"/>, in document order.</summary>
    public static List<TableRow> Rows(Table table) =>
        table.Descendants<TableRow>().Where(tr => ReferenceEquals(tr.Ancestors<Table>().First(), table)).ToList();

    /// <summary>Every <c>w:tc</c> whose nearest enclosing <c>w:tr</c> is <paramref name="row"/>, in document order.</summary>
    public static List<TableCell> Cells(TableRow row) =>
        row.Descendants<TableCell>().Where(tc => ReferenceEquals(tc.Ancestors<TableRow>().First(), row)).ToList();

    /// <summary>The cell's <c>w:gridSpan</c>, 1 when absent.</summary>
    public static int Span(TableCell cell) =>
        cell.GetFirstChild<TableCellProperties>()?.GetFirstChild<GridSpan>()?.Val?.Value ?? 1;

    /// <summary><c>gridBefore + Σ gridSpan + gridAfter</c> for <paramref name="row"/>.</summary>
    public static int Coverage(TableRow row)
    {
        var trPr = row.GetFirstChild<TableRowProperties>();
        var before = trPr?.GetFirstChild<GridBefore>()?.Val?.Value ?? 0;
        var after = trPr?.GetFirstChild<GridAfter>()?.Val?.Value ?? 0;
        return before + after + Cells(row).Sum(Span);
    }

    /// <summary>The number of <c>w:gridCol</c>s <paramref name="table"/> declares.</summary>
    public static int GridCount(Table table) =>
        table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().Count() ?? 0;

    /// <summary>
    /// One line per row, anywhere in <paramref name="doc"/>'s body, that does not cover its table's declared
    /// grid exactly — empty when every row of every table with a grid does. Phrased as strings so a failing
    /// assertion names the table, the row and both numbers.
    /// </summary>
    public static List<string> RowsOffGrid(WordprocessingDocument doc)
    {
        var body = doc.MainDocumentPart!.Document!.Body!;
        var offenders = new List<string>();
        var tables = body.Descendants<Table>().ToList();
        for (var t = 0; t < tables.Count; t++)
        {
            var grid = GridCount(tables[t]);
            if (grid == 0)
            {
                continue;
            }

            var rows = Rows(tables[t]);
            for (var r = 0; r < rows.Count; r++)
            {
                var coverage = Coverage(rows[r]);
                if (coverage != grid)
                {
                    offenders.Add($"table {t} w:tr {r} covers {coverage} of {grid} grid column(s)");
                }
            }
        }

        return offenders;
    }

    /// <summary>The nearest enclosing table of <paramref name="element"/>.</summary>
    public static Table NearestTable(OpenXmlElement element) => element.Ancestors<Table>().First();
}
