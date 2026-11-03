using BcWordLayout.Domain;
using BcWordLayout.McpHost;
using BcWordLayout.McpHost.Tools;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

/// <summary>
/// GitHub issue #9: vertical merges (<c>w:vMerge</c>). Every table tool refused them while no reviewed layout had one;
/// Microsoft's BC 29 header/footer designs do, so the tools now keep a merge whole instead (see
/// <see cref="VerticalMerges"/>).
/// </summary>
/// <remarks>
/// <para>
/// Fixtures: <see cref="Corpus.ExternalDefault"/>'s <c>header2.xml</c> table 0 (two columns; column 0, the report
/// caption, merged over rows 0–2; column 1 holds a page number, a date and the company name) and
/// <see cref="Corpus.SpendRequestDocument"/>'s table 0 (six grid columns; the <c>gridSpan</c> 2 Purpose value at grid
/// column 1 merged over rows 3–4, which also skip their last grid column).
/// </para>
/// <para>
/// Every assertion about the result reads the cells with <see cref="Walk"/>, a walk of its own, not
/// <see cref="VerticalMerges"/>, so these tests do not share a blind spot with the code they test.
/// </para>
/// </remarks>
public class VerticalMergeTests
{
    private const string Header = "header2.xml";

    private static string CopyOfCorpus(string corpusFile)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-vmerge-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    /// <summary>One cell as the file has it: grid column, span, <c>w:vMerge</c> (R, C or empty) and text.</summary>
    private sealed record CellView(int GridColumn, int Span, string Merge, string Text);

    /// <summary>The rows of table <paramref name="tableIndex"/> in <paramref name="part"/>, cell by cell.</summary>
    private static List<List<CellView>> Walk(string path, string part, int tableIndex)
    {
        using var doc = WordprocessingDocument.Open(path, false);
        OpenXmlElement root = part == "document.xml"
            ? doc.MainDocumentPart!.Document!.Body!
            : doc.MainDocumentPart!.HeaderParts.Single(h => h.Uri.OriginalString.EndsWith(part, StringComparison.Ordinal)).Header!;
        var table = root.Descendants<Table>().ElementAt(tableIndex);
        var rows = new List<List<CellView>>();
        foreach (var tr in table.Descendants<TableRow>().Where(r => ReferenceEquals(r.Ancestors<Table>().First(), table)))
        {
            var grid = tr.TableRowProperties?.GetFirstChild<GridBefore>()?.Val?.Value ?? 0;
            var cells = new List<CellView>();
            foreach (var tc in tr.Descendants<TableCell>().Where(c => ReferenceEquals(c.Ancestors<TableRow>().First(), tr)))
            {
                var span = tc.TableCellProperties?.GridSpan?.Val?.Value ?? 1;
                var merge = tc.TableCellProperties?.VerticalMerge is { } vm
                    ? vm.Val?.Value == MergedCellValues.Restart ? "R" : "C"
                    : string.Empty;
                cells.Add(new CellView(grid, span, merge, tc.InnerText));
                grid += span;
            }

            rows.Add(cells);
        }

        return rows;
    }

    private static void AssertSound(string path)
    {
        using var doc = WordprocessingDocument.Open(path, false);
        Assert.Empty(TableGridConsistencyGuard.Find(doc));
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc));
    }

    private static T Data<T>(ToolResponse response)
    {
        Assert.True(response.Ok, response.Error?.Message);
        return Assert.IsType<T>(response.Data);
    }

    private static string Refusal(ToolResponse response)
    {
        Assert.False(response.Ok);
        Assert.Equal("invalid_argument", response.Error!.Code);
        return response.Error.Message;
    }

    // ---- reading ----

    [Fact]
    public void Get_layout_info_reports_each_merged_cell()
    {
        var info = Data<LayoutInfoDto>(ReadTools.GetLayoutInfo(Corpus.Path(Corpus.ExternalDefault)));
        var header = info.Tables.Single(t => t.Part == Header && t.TableIndex == 0);
        Assert.Equal(["restart", "continue", "continue"], header.Rows.Select(r => r.Cells[0].VerticalMerge));
        Assert.All(header.Rows, r => Assert.Null(r.Cells[1].VerticalMerge));

        info = Data<LayoutInfoDto>(ReadTools.GetLayoutInfo(Corpus.Path(Corpus.SpendRequestDocument)));
        var grid = info.Tables.Single(t => t.Part == "document.xml" && t.TableIndex == 0);
        Assert.Equal("restart", grid.Rows[3].Cells[1].VerticalMerge);
        Assert.Equal("continue", grid.Rows[4].Cells[1].VerticalMerge);
        Assert.Equal(2, grid.Rows[4].Cells[1].GridSpan);
    }

    [Fact]
    public void Both_corpus_layouts_start_out_sound()
    {
        AssertSound(Corpus.Path(Corpus.ExternalDefault));
        AssertSound(Corpus.Path(Corpus.SpendRequestDocument));
    }

    // ---- columns ----

    [Fact]
    public void Set_column_widths_resizes_every_row_of_a_merge()
    {
        var path = CopyOfCorpus(Corpus.ExternalDefault);
        try
        {
            Data<TableEditResultDto>(TableTools.SetColumnWidths(path, 0, [5000, 5495], "header", Header));

            using (var doc = WordprocessingDocument.Open(path, false))
            {
                var table = doc.MainDocumentPart!.HeaderParts
                    .Single(h => h.Uri.OriginalString.EndsWith(Header, StringComparison.Ordinal)).Header!.Descendants<Table>().First();
                Assert.All(
                    table.Descendants<TableRow>(),
                    tr => Assert.Equal("5000", tr.Elements<TableCell>().First().TableCellProperties!.TableCellWidth!.Width!.Value));
            }

            Assert.Equal(["R", "C", "C"], Walk(path, Header, 0).Select(r => r[0].Merge));
            AssertSound(path);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Appending_a_column_beside_a_merged_caption_adds_a_cell_to_every_row()
    {
        // Rows 1 and 2 start with the merge's empty continue cells and end with content (the date, the company
        // name): read cell by cell, they look like a totals block and used to have their last cell widened. The
        // merge shows the report caption there, so they are not, and every row gets the new cell.
        var path = CopyOfCorpus(Corpus.ExternalDefault);
        try
        {
            var result = Data<TableEditResultDto>(TableTools.InsertColumn(path, 0, "plainText", layoutPart: "header", partName: Header));
            Assert.Equal(3, result.ColumnCountAfter);

            var rows = Walk(path, Header, 0);
            Assert.All(rows, r => Assert.Equal([0, 1, 2], r.Select(c => c.GridColumn)));
            Assert.All(rows, r => Assert.All(r, c => Assert.Equal(1, c.Span)));
            Assert.Equal(["R", "C", "C"], rows.Select(r => r[0].Merge));
            Assert.All(rows, r => Assert.Equal(string.Empty, r[2].Merge));
            AssertSound(path);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_column_inserted_left_of_a_merge_moves_it_whole()
    {
        var path = CopyOfCorpus(Corpus.ExternalDefault);
        try
        {
            Data<TableEditResultDto>(TableTools.InsertColumn(path, 0, "plainText", atColumn: 0, layoutPart: "header", partName: Header));

            var rows = Walk(path, Header, 0);
            Assert.All(rows, r => Assert.Equal(string.Empty, r[0].Merge));
            Assert.Equal(["R", "C", "C"], rows.Select(r => r[1].Merge));
            Assert.All(rows, r => Assert.Equal(1, r[1].GridColumn));
            AssertSound(path);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Appending_beside_a_merged_last_cell_widens_it_in_every_row_of_the_merge()
    {
        // The NA demo invoices' shape: a due-date column merged down the address block. Row 0 reads as a totals row
        // (empty first cell, content last) and widens its last cell to keep it on the table's edge. Row 1's first cell
        // has text, so on its own it would get a new cell instead and cut the merge; it widens with the merge.
        var path = SyntheticLayout.Create(Table(
            Row(Cell(""), Cell("From"), Cell("Due 30 days", "restart")),
            Row(Cell("Address"), Cell("Street"), Cell("", "continue")),
            Row(Cell(""), Cell("City"), Cell("", "continue"))));
        try
        {
            Data<TableEditResultDto>(TableTools.InsertColumn(path, 0, "plainText"));

            var rows = Walk(path, "document.xml", 0);
            Assert.All(rows, r => Assert.Equal(3, r.Count));
            Assert.All(rows, r => Assert.Equal((2, 2), (r[2].GridColumn, r[2].Span)));
            Assert.Equal(["R", "C", "C"], rows.Select(r => r[2].Merge));
            AssertSound(path);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Appending_is_refused_when_a_row_of_the_merge_must_take_the_new_column_content()
    {
        // Row 1 is the header row (w:tblHeader), so it gets the new header text in a new cell, while row 0 widens its
        // merged last cell: the merge would be cut in two.
        var path = SyntheticLayout.Create(Table(
            Row(Cell(""), Cell("Due 30 days", "restart")),
            Row(Cell("Caption"), Cell("", "continue"), header: true)));
        try
        {
            var message = Refusal(TableTools.InsertColumn(path, 0, "plainText", headerText: "New"));
            Assert.Contains("vertically merged", message, StringComparison.Ordinal);
            Assert.Contains("must receive the new column's content", message, StringComparison.Ordinal);
            Assert.Equal(["R", "C"], Walk(path, "document.xml", 0).Select(r => r[1].Merge));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Removing_the_merged_column_removes_every_row_of_the_merge()
    {
        var path = CopyOfCorpus(Corpus.ExternalDefault);
        try
        {
            Data<TableEditResultDto>(TableTools.RemoveColumn(path, 0, 0, "header", Header));

            var rows = Walk(path, Header, 0);
            Assert.All(rows, r => Assert.Equal(string.Empty, Assert.Single(r).Merge));
            AssertSound(path);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Removing_a_column_under_a_spanned_merge_narrows_it_in_both_rows()
    {
        var path = CopyOfCorpus(Corpus.SpendRequestDocument);
        try
        {
            Data<TableEditResultDto>(TableTools.RemoveColumn(path, 0, 2));

            var rows = Walk(path, "document.xml", 0);
            Assert.Equal(("R", 1, 1), (rows[3][1].Merge, rows[3][1].GridColumn, rows[3][1].Span));
            Assert.Equal(("C", 1, 1), (rows[4][1].Merge, rows[4][1].GridColumn, rows[4][1].Span));
            AssertSound(path);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- rows ----

    [Fact]
    public void A_row_inserted_inside_a_merge_continues_it()
    {
        var path = CopyOfCorpus(Corpus.ExternalDefault);
        try
        {
            var result = Data<TableEditResultDto>(TableTools.InsertTableRow(
                path, 0, [new RowCellArgument(), new RowCellArgument()], atRow: 1, layoutPart: "header", partName: Header));
            Assert.Contains("continues the vertically merged cell(s)", result.Summary, StringComparison.Ordinal);
            Assert.Contains("now rows 0..3", result.Summary, StringComparison.Ordinal);

            Assert.Equal(["R", "C", "C", "C"], Walk(path, Header, 0).Select(r => r[0].Merge));
            AssertSound(path);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(0)]
    [InlineData(3)]
    public void A_row_inserted_before_or_after_a_merge_stays_out_of_it(int atRow)
    {
        var path = CopyOfCorpus(Corpus.ExternalDefault);
        try
        {
            var result = Data<TableEditResultDto>(TableTools.InsertTableRow(
                path, 0, [new RowCellArgument(), new RowCellArgument()], atRow: atRow, layoutPart: "header", partName: Header));
            Assert.DoesNotContain("merged", result.Summary, StringComparison.Ordinal);

            var merges = Walk(path, Header, 0).Select(r => r[0].Merge).ToList();
            Assert.Equal(string.Empty, merges[atRow]);
            Assert.Equal(3, merges.Count(m => m.Length > 0));
            AssertSound(path);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_row_inside_a_merge_whose_cells_do_not_line_up_with_it_is_refused()
    {
        var path = CopyOfCorpus(Corpus.ExternalDefault);
        try
        {
            var message = Refusal(TableTools.InsertTableRow(
                path, 0, [new RowCellArgument { Span = 2 }], atRow: 2, layoutPart: "header", partName: Header));
            Assert.Contains("falls inside a vertically merged cell", message, StringComparison.Ordinal);
            Assert.Contains("use atRow 0 (just before the merge) or 3 (just after it)", message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_row_inside_a_merge_that_puts_content_under_it_is_refused()
    {
        var path = CopyOfCorpus(Corpus.ExternalDefault);
        try
        {
            var message = Refusal(TableTools.InsertTableRow(
                path,
                0,
                [new RowCellArgument { Fields = ["/BCReportInformation/ReportRequest/DateAndTime"] }, new RowCellArgument()],
                atRow: 1,
                layoutPart: "header",
                partName: Header));
            Assert.Contains("Word would hide under the merged cell's content", message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- cells ----

    [Fact]
    public void Merge_cells_refuses_to_cut_through_a_vertical_merge()
    {
        var path = CopyOfCorpus(Corpus.ExternalDefault);
        try
        {
            var message = Refusal(TableTools.MergeCells(path, 0, row: 1, fromColumn: 0, toColumn: 1, "header", Header));
            Assert.Contains("vertically merged", message, StringComparison.Ordinal);
            Assert.Contains("rows 0..2", message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Split_cells_splits_every_row_of_a_merge_and_keeps_each_column_merged()
    {
        var path = CopyOfCorpus(Corpus.SpendRequestDocument);
        try
        {
            var result = Data<TableEditResultDto>(TableTools.SplitCells(path, 0, row: 3, cellIndex: 1));
            Assert.Contains("in every row of its vertical merge (rows 3..4)", result.Summary, StringComparison.Ordinal);

            var rows = Walk(path, "document.xml", 0);
            Assert.Equal([("R", 1), ("R", 2)], rows[3].Skip(1).Take(2).Select(c => (c.Merge, c.GridColumn)));
            Assert.Equal([("C", 1), ("C", 2)], rows[4].Skip(1).Take(2).Select(c => (c.Merge, c.GridColumn)));
            AssertSound(path);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Text_for_a_continue_cell_is_refused_naming_the_merge_first_cell()
    {
        var path = CopyOfCorpus(Corpus.ExternalDefault);
        try
        {
            var message = Refusal(EditTools.SetCellText(path, 0, row: 2, col: 0, "x", "header", Header));
            Assert.Contains("Address that cell instead: row 0, col 0", message, StringComparison.Ordinal);

            var insert = EditTools.InsertText(path, "x", "tableCell", tableIndex: 0, row: 1, col: 0, layoutPart: "header", partName: Header);
            Assert.Contains("rows 0..2", Refusal(insert), StringComparison.Ordinal);
            Assert.Contains("verticalMerge 'restart'", insert.Error!.Hint, StringComparison.Ordinal);

            // Clearing one is harmless: there is nothing in it Word would show.
            Data<CellEditResultDto>(EditTools.ClearCellText(path, 0, row: 2, col: 0, "header", Header));
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- the guard ----

    [Fact]
    public void The_guard_reports_a_continue_cell_with_nothing_above_it_to_continue()
    {
        var path = SyntheticLayout.Create(Table(
            Row(Cell("a"), Cell("b")),
            Row(Cell("", "continue"), Cell("c"))));
        try
        {
            using var doc = WordprocessingDocument.Open(path, false);
            var violation = Assert.Single(TableGridConsistencyGuard.Find(doc));
            Assert.Equal(TableGridViolationKind.BrokenVerticalMerge, violation.Kind);
            Assert.Equal(1, violation.RowIndex);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- synthetic tables ----

    private static string Table(params string[] rows)
    {
        var columns = rows[0].Split("<w:tc>").Length - 1;
        return "<w:tbl><w:tblPr/><w:tblGrid>"
            + string.Concat(Enumerable.Repeat("<w:gridCol w:w=\"2000\"/>", columns))
            + "</w:tblGrid>" + string.Concat(rows) + "</w:tbl>";
    }

    private static string Row(string a, string b, string? c = null, bool header = false) =>
        "<w:tr>" + (header ? "<w:trPr><w:tblHeader/></w:trPr>" : string.Empty) + a + b + c + "</w:tr>";

    private static string Row(string a, string b, bool header) => Row(a, b, null, header);

    private static string Cell(string text, string? merge = null) =>
        "<w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/>"
        + merge switch
        {
            "restart" => "<w:vMerge w:val=\"restart\"/>",
            "continue" => "<w:vMerge/>",
            _ => string.Empty,
        }
        + "</w:tcPr><w:p>" + (text.Length == 0 ? string.Empty : $"<w:r><w:t>{text}</w:t></w:r>") + "</w:p></w:tc>";
}
