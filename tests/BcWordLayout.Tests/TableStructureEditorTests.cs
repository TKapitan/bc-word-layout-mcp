using System.Globalization;
using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;
using BcWordLayout.McpHost;
using BcWordLayout.McpHost.Tools;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

/// <summary>
/// Covers <see cref="TableStructureEditor"/> (set_column_widths / insert_column / remove_column) and its
/// <see cref="TableGridConsistencyGuard"/> backstop, end to end against COPIES of the real corpus
/// (SalesInvoice's <c>/Header/Line</c> table — a 9-grid-column, pervasively gridSpan-spanned, ragged
/// line-items table, the flagship P6/F10 target) plus small synthetic tables for precise assertions.
/// </summary>
public class TableStructureEditorTests
{
    // The SalesInvoice body's Line table is the 3rd table in document order (index 2), with a 9-column grid.
    private const int LineTableIndex = 2;
    private const int LineGridColumns = 9;
    private const string TaxAmountXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line[1]/ns0:TransHeaderAmount[1]";

    private static string CopyOfCorpus(string corpusFile)
    {
        var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"bcwl-tablestruct-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    private static List<ValidationErrorInfo> OpenXmlErrors(WordprocessingDocument doc) =>
        new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc).ToList();

    private static Table BodyTable(WordprocessingDocument doc, int index) =>
        doc.MainDocumentPart!.Document!.Body!.Descendants<Table>().ElementAt(index);

    private static List<string> GridWidths(Table t) =>
        t.GetFirstChild<TableGrid>()!.Elements<GridColumn>().Select(g => g.Width!.Value!).ToList();

    // ---- the guard does not false-flag real tables ----

    [Theory]
    [InlineData(Corpus.SalesInvoice)]
    [InlineData(Corpus.InventoryOrderDetails)]
    [InlineData(Corpus.StandardStatement)]
    public void Guard_reports_no_violations_on_pristine_corpus(string corpusFile)
    {
        using var doc = WordprocessingDocument.Open(Corpus.Path(corpusFile), false);
        var violations = TableGridConsistencyGuard.Find(doc);
        Assert.True(violations.Count == 0,
            "unexpected grid violations: " + string.Join(" | ", violations.Select(v => v.Describe())));
    }

    [Fact]
    public void Guard_flags_a_row_that_does_not_cover_its_grid()
    {
        // grid declares 3 columns but the single row has only 2 unit cells.
        var body =
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"1000\"/><w:gridCol w:w=\"1000\"/><w:gridCol w:w=\"1000\"/></w:tblGrid>"
            + "<w:tr><w:tc><w:tcPr/><w:p/></w:tc><w:tc><w:tcPr/><w:p/></w:tc></w:tr></w:tbl>";
        var path = SyntheticLayout.Create(body);
        try
        {
            using var doc = WordprocessingDocument.Open(path, false);
            var violations = TableGridConsistencyGuard.Find(doc);
            Assert.Single(violations);
            Assert.Contains("cover 2 grid column(s)", violations[0].Reason);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- set_column_widths ----

    [Fact]
    public void SetColumnWidths_on_corpus_line_table_sets_grid_and_keeps_it_consistent()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var widths = new[] { 1000, 2000, 800, 500, 500, 400, 1300, 700, 1600 };
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var result = TableStructureEditor.SetColumnWidths(doc, LayoutPart.Body, null, LineTableIndex, widths);
                Assert.Equal("set_column_widths", result.Operation);
                Assert.Equal(LineGridColumns, result.ColumnCountBefore);
                Assert.Equal(LineGridColumns, result.ColumnCountAfter);
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            Assert.Equal(widths.Select(w => w.ToString()), GridWidths(BodyTable(reopened, LineTableIndex)));
            Assert.Empty(TableGridConsistencyGuard.Find(reopened));
            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void SetColumnWidths_wrong_count_throws_ArgumentException()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var ex = Assert.Throws<ArgumentException>(() =>
                TableStructureEditor.SetColumnWidths(doc, LayoutPart.Body, null, LineTableIndex, new[] { 1, 2, 3 }));
            Assert.Equal("widths", ex.ParamName);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void SetColumnWidths_on_simple_table_sets_every_cell_width()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(rows: 3, cols: 3));
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                TableStructureEditor.SetColumnWidths(doc, LayoutPart.Body, null, 0, new[] { 1500, 2500, 3500 });
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var table = BodyTable(reopened, 0);
            Assert.Equal(new[] { "1500", "2500", "3500" }, GridWidths(table));
            foreach (var row in table.Elements<TableRow>())
            {
                var widths = row.Elements<TableCell>()
                    .Select(c => c.GetFirstChild<TableCellProperties>()?.GetFirstChild<TableCellWidth>()?.Width?.Value)
                    .ToList();
                Assert.Equal(new[] { "1500", "2500", "3500" }, widths);
            }

            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void SetColumnWidths_writes_and_reads_invariant_digits_regardless_of_current_culture()
    {
        // w:gridCol/@w:w is an invariant OOXML numeric attribute. Under some cultures (ar-SA, fa-IR),
        // int.ToString(CurrentCulture) renders the negative sign as a different, non-ASCII character, and
        // the 2-arg int.TryParse(string) overload FAILS to parse a perfectly normal ASCII "-100" (verified:
        // both cultures reject it under CurrentCulture, uncorrupted by this test). Pinning every write/read
        // of these values to InvariantCulture keeps the document's own numeric attributes ASCII always, and
        // keeps reads working regardless of the host thread's culture.
        var original = CultureInfo.CurrentCulture;
        CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("ar-SA");
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(rows: 1, cols: 1));
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                TableStructureEditor.SetColumnWidths(doc, LayoutPart.Body, null, 0, new[] { -100 });
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var table = BodyTable(reopened, 0);

            // The raw XML attribute must be the plain invariant ASCII string, not ar-SA's mangled
            // negative sign (which int.ToString(CurrentCulture) would have produced without the fix).
            Assert.Equal("-100", table.GetFirstChild<TableGrid>()!.GetFirstChild<GridColumn>()!.Width!.Value);

            // Reading it back (TableStructureReader.GridColumnWidths via LayoutReader) must not silently
            // drop it either, regardless of CurrentCulture at read time (still ar-SA here).
            var inv = LayoutReader.Read(path);
            Assert.Equal(-100, Assert.Single(inv.Tables).GridColumnWidths.Single());
        }
        finally
        {
            CultureInfo.CurrentCulture = original;
            File.Delete(path);
        }
    }

    // ---- insert_column ----

    [Fact]
    public void InsertColumn_append_field_on_corpus_line_table_adds_grid_column_and_bound_fields_to_every_repeater_row()
    {
        // The real SalesInvoice Line table (index 2) carries TWO row-level repeaters sharing the same
        // physical w:tbl - the Line item row itself AND the ReportTotalsLine row a few rows below it (both
        // are "control rows" per TableGridNavigator). insert_column's per-row rule binds the new field to
        // EVERY control row, not just one, so the grid stays consistent across both repeaters: 2 new bound
        // cells, not 1.
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            int fieldsBefore;
            using (var probe = WordprocessingDocument.Open(path, false))
            {
                fieldsBefore = LayoutReader.Read(probe).Controls.Count(c => c.XPath == TaxAmountXPath);
            }

            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var result = TableStructureEditor.InsertColumn(
                    doc, LayoutPart.Body, null, LineTableIndex, atColumn: null,
                    new InsertColumnOptions { Mode = InsertColumnMode.Field, DataPath = "/Header/Line/TransHeaderAmount" });

                Assert.Equal("insert_column", result.Operation);
                Assert.Equal(LineGridColumns, result.ColumnCountBefore);
                Assert.Equal(LineGridColumns + 1, result.ColumnCountAfter);
                Assert.Equal(LineGridColumns, result.ColumnIndex); // new column's 0-based grid index
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            Assert.Equal(LineGridColumns + 1, BodyTable(reopened, LineTableIndex).GetFirstChild<TableGrid>()!.Elements<GridColumn>().Count());
            Assert.Equal(fieldsBefore + 2, LayoutReader.Read(reopened).Controls.Count(c => c.XPath == TaxAmountXPath));
            Assert.Empty(TableGridConsistencyGuard.Find(reopened));
            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- interior (non-append) positions ----

    [Fact]
    public void InsertColumn_at_an_interior_boundary_inserts_the_new_cell_in_position_in_every_row()
    {
        // Grid column 3 is a cell boundary in every row of the corpus Line table (its spanned cells all sit
        // further right), so the new column slots straight in — the bound field lands in the repeater data
        // rows, and every row's coverage still matches the grid.
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            int fieldsBefore;
            List<string> widthsBefore;
            using (var probe = WordprocessingDocument.Open(path, false))
            {
                fieldsBefore = LayoutReader.Read(probe).Controls.Count(c => c.XPath == TaxAmountXPath);
                widthsBefore = GridWidths(BodyTable(probe, LineTableIndex));
            }

            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var result = TableStructureEditor.InsertColumn(
                    doc, LayoutPart.Body, null, LineTableIndex, atColumn: 3,
                    new InsertColumnOptions
                    {
                        Mode = InsertColumnMode.Field,
                        DataPath = "/Header/Line/TransHeaderAmount",
                        Width = 700,
                    });

                Assert.Equal(3, result.ColumnIndex);
                Assert.Equal(LineGridColumns + 1, result.ColumnCountAfter);
                Assert.Contains("at grid index 3", result.Summary, StringComparison.Ordinal);
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var table = BodyTable(reopened, LineTableIndex);

            // The new w:gridCol sits AT position 3, not at the end.
            var widthsAfter = GridWidths(table);
            Assert.Equal(LineGridColumns + 1, widthsAfter.Count);
            Assert.Equal("700", widthsAfter[3]);
            Assert.Equal(widthsBefore.Take(3), widthsAfter.Take(3));
            Assert.Equal(widthsBefore.Skip(3), widthsAfter.Skip(4));

            Assert.Equal(fieldsBefore + 2, LayoutReader.Read(reopened).Controls.Count(c => c.XPath == TaxAmountXPath));
            Assert.Empty(TableGridConsistencyGuard.Find(reopened));
            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertColumn_interior_inside_a_spanned_CONTENT_cell_is_refused_naming_the_cell()
    {
        // A header row whose first cell spans grid 0..1: there is no honest place for the new header cell
        // at grid 1 — widening would drop the header text, splitting would rewrite a layout decision the
        // caller never mentioned.
        var path = SyntheticLayout.Create(SpannedHeaderTable);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var ex = Assert.Throws<ArgumentException>(() => TableStructureEditor.InsertColumn(
                doc, LayoutPart.Body, null, 0, atColumn: 1,
                new InsertColumnOptions { Mode = InsertColumnMode.PlainText, HeaderText = "Notes" }));

            Assert.Equal("atColumn", ex.ParamName);
            Assert.Contains("spans grid columns 0..1", ex.Message, StringComparison.Ordinal);
            Assert.Contains("split_cells", ex.Message, StringComparison.Ordinal);

            // Refused before touching anything.
            Assert.Equal(3, TableGridNavigator.GridColumnCount(BodyTable(doc, 0)));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertColumn_interior_inside_a_spanned_FILLER_cell_widens_it_instead_of_adding_a_cell()
    {
        // Same table, but the straddling row (a plain spacer with no content) simply gets wider: no cell is
        // added there, so nothing shifts, and its w:tcW grows by the new column's width.
        var path = SyntheticLayout.Create(SpannedFillerTable);
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var result = TableStructureEditor.InsertColumn(
                    doc, LayoutPart.Body, null, 0, atColumn: 1,
                    new InsertColumnOptions { Mode = InsertColumnMode.PlainText, HeaderText = "Notes", Width = 500 });

                Assert.Equal(4, result.ColumnCountAfter);
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var table = BodyTable(reopened, 0);
            var rows = table.Elements<TableRow>().ToList();

            // Row 0 (unit cells) gained a cell at index 1; row 1's spanning cell absorbed the column.
            Assert.Equal(4, rows[0].Elements<TableCell>().Count());
            var spannedRowCells = rows[1].Elements<TableCell>().ToList();
            Assert.Equal(2, spannedRowCells.Count);
            Assert.Equal(3, spannedRowCells[0].TableCellProperties!.GridSpan!.Val!.Value);
            Assert.Equal("2500", spannedRowCells[0].TableCellProperties!.TableCellWidth!.Width!.Value); // 1000+500+1000

            Assert.Empty(TableGridConsistencyGuard.Find(reopened));
            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    /// <summary>
    /// A 3-column table whose HEADER row (<c>w:tblHeader</c>) starts with a cell spanning grid 0..1 — the
    /// shape an interior insert at grid 1 must refuse (that row has to carry the new header cell).
    /// </summary>
    internal const string SpannedHeaderTable =
        "<w:tbl><w:tblPr/><w:tblGrid>"
        + "<w:gridCol w:w=\"1000\"/><w:gridCol w:w=\"1000\"/><w:gridCol w:w=\"1000\"/></w:tblGrid>"
        + "<w:tr><w:trPr><w:tblHeader/></w:trPr>"
        + "<w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/><w:gridSpan w:val=\"2\"/></w:tcPr>"
        + "<w:p><w:r><w:t>Item</w:t></w:r></w:p></w:tc>"
        + "<w:tc><w:tcPr><w:tcW w:w=\"1000\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>Amount</w:t></w:r></w:p></w:tc>"
        + "</w:tr>"
        + "<w:tr>"
        + "<w:tc><w:tcPr><w:tcW w:w=\"1000\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc>"
        + "<w:tc><w:tcPr><w:tcW w:w=\"1000\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc>"
        + "<w:tc><w:tcPr><w:tcW w:w=\"1000\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>c</w:t></w:r></w:p></w:tc>"
        + "</w:tr></w:tbl>";

    /// <summary>
    /// A 3-column table whose second row is an EMPTY spacer spanning grid 0..1 — the straddling row an
    /// interior insert at grid 1 widens rather than refusing (no content of its own to place).
    /// </summary>
    private const string SpannedFillerTable =
        "<w:tbl><w:tblPr/><w:tblGrid>"
        + "<w:gridCol w:w=\"1000\"/><w:gridCol w:w=\"1000\"/><w:gridCol w:w=\"1000\"/></w:tblGrid>"
        + "<w:tr><w:trPr><w:tblHeader/></w:trPr>"
        + "<w:tc><w:tcPr><w:tcW w:w=\"1000\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>Item</w:t></w:r></w:p></w:tc>"
        + "<w:tc><w:tcPr><w:tcW w:w=\"1000\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>Qty</w:t></w:r></w:p></w:tc>"
        + "<w:tc><w:tcPr><w:tcW w:w=\"1000\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>Amount</w:t></w:r></w:p></w:tc>"
        + "</w:tr>"
        + "<w:tr>"
        + "<w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/><w:gridSpan w:val=\"2\"/></w:tcPr><w:p/></w:tc>"
        + "<w:tc><w:tcPr><w:tcW w:w=\"1000\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>822.97</w:t></w:r></w:p></w:tc>"
        + "</w:tr></w:tbl>";

    [Fact]
    public void InsertColumn_field_mode_without_dataPath_is_rejected()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var ex = Assert.Throws<ArgumentException>(() => TableStructureEditor.InsertColumn(
                doc, LayoutPart.Body, null, LineTableIndex, atColumn: null,
                new InsertColumnOptions { Mode = InsertColumnMode.Field, DataPath = null }));
            Assert.Equal("dataPath", ex.ParamName);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertColumn_plainText_on_simple_table_appends_a_cell_per_row()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(rows: 3, cols: 3));
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var result = TableStructureEditor.InsertColumn(
                    doc, LayoutPart.Body, null, 0, atColumn: null,
                    new InsertColumnOptions { Mode = InsertColumnMode.PlainText, HeaderText = "Notes" });
                Assert.Equal(4, result.ColumnCountAfter);
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var table = BodyTable(reopened, 0);
            Assert.Equal(4, table.GetFirstChild<TableGrid>()!.Elements<GridColumn>().Count());
            Assert.All(table.Elements<TableRow>(), r => Assert.Equal(4, r.Elements<TableCell>().Count()));
            // Row 0 (no tblHeader, no repeater) is treated as the header only for repeater tables; for a plain
            // table the header text lands only where a header row exists — here none, so "Notes" is not forced.
            Assert.Empty(TableGridConsistencyGuard.Find(reopened));
            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertColumn_new_cells_inherit_the_neighbor_cells_formatting()
    {
        // Mirrors the real BC header-row look: per-cell bottom border + bottom vAlign on the cells,
        // right-justified paragraphs, bold runs. Without cloning, an appended column visibly does not
        // belong to the table (top-aligned unbordered header next to bottom-aligned underlined ones -
        // the exact defect the 2026-07-31 edit-scenario e2e pass caught on StandardPurchaseOrder).
        const string formattedCell =
            "<w:tc><w:tcPr><w:tcBorders><w:bottom w:val=\"single\"/></w:tcBorders><w:vAlign w:val=\"bottom\"/></w:tcPr>"
            + "<w:p><w:pPr><w:jc w:val=\"right\"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Amount</w:t></w:r></w:p></w:tc>";
        var body =
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"3000\"/><w:gridCol w:w=\"3000\"/></w:tblGrid>"
            + $"<w:tr><w:trPr><w:tblHeader/></w:trPr>{formattedCell}{formattedCell}</w:tr>"
            + $"<w:tr>{formattedCell}{formattedCell}</w:tr></w:tbl>";
        var path = SyntheticLayout.Create(body);
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                TableStructureEditor.InsertColumn(
                    doc, LayoutPart.Body, null, 0, atColumn: null,
                    new InsertColumnOptions { Mode = InsertColumnMode.PlainText, HeaderText = "Notes" });
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var rows = BodyTable(reopened, 0).Elements<TableRow>().ToList();

            var headerCell = rows[0].Elements<TableCell>().Last();
            Assert.Equal("Notes", headerCell.InnerText);
            var headerTcPr = headerCell.GetFirstChild<TableCellProperties>()!;
            Assert.NotNull(headerTcPr.GetFirstChild<TableCellBorders>()?.BottomBorder);
            Assert.Equal(TableVerticalAlignmentValues.Bottom, headerTcPr.GetFirstChild<TableCellVerticalAlignment>()!.Val!.Value);
            var headerParagraph = headerCell.GetFirstChild<Paragraph>()!;
            Assert.Equal(JustificationValues.Right, headerParagraph.ParagraphProperties!.Justification!.Val!.Value);
            Assert.NotNull(headerParagraph.GetFirstChild<Run>()!.RunProperties?.Bold);

            var dataCell = rows[1].Elements<TableCell>().Last();
            Assert.Equal(string.Empty, dataCell.InnerText);
            Assert.NotNull(dataCell.GetFirstChild<TableCellProperties>()!.GetFirstChild<TableCellBorders>()?.BottomBorder);
            Assert.Equal(
                JustificationValues.Right,
                dataCell.GetFirstChild<Paragraph>()!.ParagraphProperties!.Justification!.Val!.Value);

            Assert.Empty(TableGridConsistencyGuard.Find(reopened));
            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertColumn_keeps_a_right_anchored_totals_row_at_the_tables_right_edge()
    {
        // The BC totals-block shape below a lines table: leading EMPTY cells, then "Total | 822.97"
        // hugging the right edge. Appending a new empty cell at such a row's END pushes the summary one
        // column in from the edge and extends its horizontal rules (cloned borders) under the new column
        // (2026-07-31 user finding on s05-po-add-two-columns-refit). Such a row gets NO new cell at all:
        // its trailing content cell is widened (gridSpan +1) across the new column, so every cell keeps
        // its position/width and the right-aligned amount stays on the table's right edge.
        var body =
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2000\"/><w:gridCol w:w=\"2000\"/><w:gridCol w:w=\"2000\"/></w:tblGrid>"
            // row 0: an ordinary full row (first cell has content) - the new cell must still be APPENDED.
            + "<w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>"
            + "<w:tc><w:tcPr/><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>"
            + "<w:tc><w:tcPr/><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc></w:tr>"
            // row 1: right-anchored totals shape - [empty][Total|bordered][822.97|bordered].
            + "<w:tr><w:tc><w:tcPr/><w:p/></w:tc>"
            + "<w:tc><w:tcPr><w:tcBorders><w:top w:val=\"single\"/></w:tcBorders></w:tcPr>"
            + "<w:p><w:r><w:t>Total</w:t></w:r></w:p></w:tc>"
            + "<w:tc><w:tcPr><w:tcBorders><w:top w:val=\"single\"/></w:tcBorders></w:tcPr>"
            + "<w:p><w:r><w:t>822.97</w:t></w:r></w:p></w:tc></w:tr></w:tbl>";
        var path = SyntheticLayout.Create(body);
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                TableStructureEditor.InsertColumn(
                    doc, LayoutPart.Body, null, 0, atColumn: null,
                    new InsertColumnOptions { Mode = InsertColumnMode.PlainText, Width = 1000 });
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var rows = BodyTable(reopened, 0).Elements<TableRow>().ToList();

            // Ordinary row: new empty cell appended at the END.
            var row0 = rows[0].Elements<TableCell>().ToList();
            Assert.Equal(4, row0.Count);
            Assert.Equal(string.Empty, row0[^1].InnerText);
            Assert.Equal("A", row0[0].InnerText);

            // Totals row: NO new cell; the amount cell widened across the new column.
            var row1 = rows[1].Elements<TableCell>().ToList();
            Assert.Equal(3, row1.Count);
            Assert.Equal(string.Empty, row1[0].InnerText);
            Assert.Equal("Total", row1[1].InnerText);
            Assert.Equal("822.97", row1[2].InnerText);
            var amountTcPr = row1[2].GetFirstChild<TableCellProperties>()!;
            Assert.Equal(2, amountTcPr.GetFirstChild<GridSpan>()!.Val!.Value);
            // Re-sliced widths: untouched cells keep their column's width; the widened amount cell covers
            // its old 2000 column plus the new 1000 one.
            Assert.Equal(
                new[] { "2000", "2000", "3000" },
                row1.Select(c => c.GetFirstChild<TableCellProperties>()!.GetFirstChild<TableCellWidth>()!.Width!.Value));

            Assert.Empty(TableGridConsistencyGuard.Find(reopened));
            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- remove_column ----

    [Fact]
    public void RemoveColumn_on_corpus_line_table_shrinks_grid_and_keeps_it_consistent()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            int widthBefore;
            using (var probe = WordprocessingDocument.Open(path, false))
            {
                widthBefore = GridWidths(BodyTable(probe, LineTableIndex)).Sum(int.Parse);
            }

            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var result = TableStructureEditor.RemoveColumn(doc, LayoutPart.Body, null, LineTableIndex, gridColumn: 6);
                Assert.Equal("remove_column", result.Operation);
                Assert.Equal(LineGridColumns, result.ColumnCountBefore);
                Assert.Equal(LineGridColumns - 1, result.ColumnCountAfter);
                Assert.True(result.RowsAffected > 0);
                Assert.Contains("redistributed", result.Summary);
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var table = BodyTable(reopened, LineTableIndex);
            Assert.Equal(LineGridColumns - 1, table.GetFirstChild<TableGrid>()!.Elements<GridColumn>().Count());
            // A full-width BC line table must still span the same total width after the removal (the
            // removed column's twips are redistributed, not silently dropped from the table's right edge).
            Assert.Equal(widthBefore, GridWidths(table).Sum(int.Parse));
            Assert.Empty(TableGridConsistencyGuard.Find(reopened));
            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RemoveColumn_redistributes_the_removed_width_across_the_remaining_columns()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(rows: 2, cols: 3)); // 3 x 2000 twips
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                TableStructureEditor.RemoveColumn(doc, LayoutPart.Body, null, 0, gridColumn: 1);
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var table = BodyTable(reopened, 0);
            Assert.Equal(new[] { "3000", "3000" }, GridWidths(table));
            foreach (var cell in table.Descendants<TableCell>())
            {
                Assert.Equal("3000", cell.GetFirstChild<TableCellProperties>()!.GetFirstChild<TableCellWidth>()!.Width!.Value);
            }

            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RemoveColumn_width_redistribution_rounds_to_the_exact_original_total()
    {
        // 1000+999+1001 = 3000; removing column 0 scales [999, 1001] by 1.5 -> 1498.5/1501.5, which no
        // per-column rounding rule hits exactly - the largest-remainder step must close the gap so the
        // table width comes out at precisely 3000, not 2999 or 3001.
        var body =
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"1000\"/><w:gridCol w:w=\"999\"/><w:gridCol w:w=\"1001\"/></w:tblGrid>"
            + "<w:tr><w:tc><w:tcPr/><w:p/></w:tc><w:tc><w:tcPr/><w:p/></w:tc><w:tc><w:tcPr/><w:p/></w:tc></w:tr></w:tbl>";
        var path = SyntheticLayout.Create(body);
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                TableStructureEditor.RemoveColumn(doc, LayoutPart.Body, null, 0, gridColumn: 0);
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var widths = GridWidths(BodyTable(reopened, 0)).Select(int.Parse).ToList();
            Assert.Equal(3000, widths.Sum());
            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RemoveColumn_on_simple_table_deletes_one_cell_per_row()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(rows: 3, cols: 3));
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                TableStructureEditor.RemoveColumn(doc, LayoutPart.Body, null, 0, gridColumn: 1);
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var table = BodyTable(reopened, 0);
            Assert.Equal(2, table.GetFirstChild<TableGrid>()!.Elements<GridColumn>().Count());
            Assert.All(table.Elements<TableRow>(), r => Assert.Equal(2, r.Elements<TableCell>().Count()));
            // The middle column (R{r}C1) is gone; R{r}C0 and R{r}C2 remain.
            Assert.DoesNotContain(table.Descendants<Text>(), t => t.Text == "R0C1");
            Assert.Contains(table.Descendants<Text>(), t => t.Text == "R0C0");
            Assert.Contains(table.Descendants<Text>(), t => t.Text == "R0C2");
            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RemoveColumn_last_remaining_column_is_rejected()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(rows: 2, cols: 1));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var ex = Assert.Throws<ArgumentException>(() =>
                TableStructureEditor.RemoveColumn(doc, LayoutPart.Body, null, 0, gridColumn: 0));
            Assert.Equal("gridColumn", ex.ParamName);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RemoveColumn_out_of_range_is_rejected()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(rows: 2, cols: 3));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var ex = Assert.Throws<ArgumentException>(() =>
                TableStructureEditor.RemoveColumn(doc, LayoutPart.Body, null, 0, gridColumn: 5));
            Assert.Equal("gridColumn", ex.ParamName);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- unsupported shapes (vMerge) are rejected, not mishandled ----

    [Fact]
    public void Column_ops_reject_a_table_that_uses_vMerge()
    {
        var body =
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2000\"/><w:gridCol w:w=\"2000\"/></w:tblGrid>"
            + "<w:tr><w:tc><w:tcPr><w:vMerge w:val=\"restart\"/></w:tcPr><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc>"
            + "<w:tc><w:tcPr/><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc></w:tr>"
            + "<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>"
            + "<w:tc><w:tcPr/><w:p><w:r><w:t>c</w:t></w:r></w:p></w:tc></w:tr></w:tbl>";
        var path = SyntheticLayout.Create(body);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var ex = Assert.Throws<ArgumentException>(() =>
                TableStructureEditor.SetColumnWidths(doc, LayoutPart.Body, null, 0, new[] { 2000, 2000 }));
            Assert.Contains("vMerge", ex.Message);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- merge_cells / split_cells (horizontal) ----

    [Fact]
    public void MergeCells_combines_adjacent_cells_into_a_gridSpan_and_keeps_the_grid()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(rows: 2, cols: 3));
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var result = TableStructureEditor.MergeCells(doc, LayoutPart.Body, null, 0, row: 0, fromColumn: 0, toColumn: 1);
                Assert.Equal("merge_cells", result.Operation);
                Assert.Equal(3, result.ColumnCountAfter); // grid unchanged
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var table = BodyTable(reopened, 0);
            var rows = table.Elements<TableRow>().ToList();
            var row0Cells = rows[0].Elements<TableCell>().ToList();
            Assert.Equal(2, row0Cells.Count); // two cells merged into one -> 2 physical cells remain
            Assert.Equal(2, row0Cells[0].GetFirstChild<TableCellProperties>()!.GetFirstChild<GridSpan>()!.Val!.Value);
            Assert.Equal(3, rows[1].Elements<TableCell>().Count()); // other row untouched
            Assert.Equal(3, table.GetFirstChild<TableGrid>()!.Elements<GridColumn>().Count());
            Assert.Empty(TableGridConsistencyGuard.Find(reopened));
            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void MergeCells_on_cells_that_already_carry_a_width_keeps_tcPr_schema_order()
    {
        // Every real BC table cell carries a w:tcW; the merged cell's new w:gridSpan must land AFTER it
        // (CT_TcPr sequence: cnfStyle, tcW, gridSpan, ...). Inserting gridSpan at the front of such a
        // tcPr is a structural error the pre-save validator gate rejects - the exact failure the
        // 2026-07-31 scenario e2e hit merging the corpus quote's spacer row (SimpleTable's cells carry
        // no tcW, which is why the sibling test above never caught it).
        const string cell = "<w:tc><w:tcPr><w:tcW w:w=\"1000\" w:type=\"dxa\"/></w:tcPr><w:p/></w:tc>";
        var body =
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"1000\"/><w:gridCol w:w=\"1000\"/><w:gridCol w:w=\"1000\"/></w:tblGrid>"
            + $"<w:tr>{cell}{cell}{cell}</w:tr></w:tbl>";
        var path = SyntheticLayout.Create(body);
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                TableStructureEditor.MergeCells(doc, LayoutPart.Body, null, 0, row: 0, fromColumn: 0, toColumn: 2);
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var merged = BodyTable(reopened, 0).Elements<TableRow>().First().Elements<TableCell>().First();
            var tcPr = merged.GetFirstChild<TableCellProperties>()!;
            Assert.Equal(3, tcPr.GetFirstChild<GridSpan>()!.Val!.Value);
            Assert.Equal("3000", tcPr.GetFirstChild<TableCellWidth>()!.Width!.Value); // widened to the span
            Assert.True(tcPr.GetFirstChild<TableCellWidth>()!.IsBefore(tcPr.GetFirstChild<GridSpan>()!),
                "w:tcW must precede w:gridSpan in tcPr");
            Assert.Empty(TableGridConsistencyGuard.Find(reopened));
            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void MergeCells_rejects_absorbing_a_cell_that_holds_a_control()
    {
        var body =
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"1000\"/><w:gridCol w:w=\"1000\"/></w:tblGrid>"
            + "<w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc>"
            + "<w:tc><w:tcPr/><w:sdt><w:sdtPr><w:id w:val=\"5\"/></w:sdtPr>"
            + "<w:sdtContent><w:p><w:r><w:t>b</w:t></w:r></w:p></w:sdtContent></w:sdt></w:tc></w:tr></w:tbl>";
        var path = SyntheticLayout.Create(body);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var ex = Assert.Throws<ArgumentException>(() =>
                TableStructureEditor.MergeCells(doc, LayoutPart.Body, null, 0, row: 0, fromColumn: 0, toColumn: 1));
            Assert.Equal("toColumn", ex.ParamName);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void SplitCell_expands_a_spanned_cell_back_to_single_columns()
    {
        var body =
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"1000\"/><w:gridCol w:w=\"1000\"/><w:gridCol w:w=\"1000\"/></w:tblGrid>"
            + "<w:tr><w:tc><w:tcPr><w:gridSpan w:val=\"2\"/></w:tcPr><w:p><w:r><w:t>AB</w:t></w:r></w:p></w:tc>"
            + "<w:tc><w:tcPr/><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc></w:tr></w:tbl>";
        var path = SyntheticLayout.Create(body);
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var result = TableStructureEditor.SplitCell(doc, LayoutPart.Body, null, 0, row: 0, cellIndex: 0);
                Assert.Equal("split_cells", result.Operation);
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var row0 = BodyTable(reopened, 0).Elements<TableRow>().First();
            var row0Cells = row0.Elements<TableCell>().ToList();
            Assert.Equal(3, row0Cells.Count); // span-2 cell split into 2 -> 3 physical cells total
            Assert.All(row0Cells, c =>
                Assert.Null(c.GetFirstChild<TableCellProperties>()?.GetFirstChild<GridSpan>()));
            Assert.Empty(TableGridConsistencyGuard.Find(reopened));
            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void SplitCell_on_an_unspanned_cell_is_rejected()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(rows: 1, cols: 3));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var ex = Assert.Throws<ArgumentException>(() =>
                TableStructureEditor.SplitCell(doc, LayoutPart.Body, null, 0, row: 0, cellIndex: 0));
            Assert.Equal("cellIndex", ex.ParamName);
        }
        finally
        {
            File.Delete(path);
        }
    }
    // ---- w:gridBefore/w:gridAfter rows (GitHub issue #9's history): supported, not rejected ----
    //
    // These rows cover fewer physical cells than the table has grid columns, so their first cell starts at
    // grid column `Before` and a target column may land in a run no cell reaches. Every operation below used
    // to refuse such a table outright on "absent from the corpus" grounds, which was wrong: gridAfter is on
    // the LINE-ITEMS table of the stock StandardSalesInvoiceVatSpec.docx layout.

    private static (int Before, int SpanTotal, int After) Coverage(Table table, int rowIndex)
    {
        var row = TableGridNavigator.Rows(table)[rowIndex].InnerRow!;
        var c = TableGridNavigator.Coverage(row);
        return (c.Before, c.SpanTotal, c.After);
    }

    private static List<string?> CellWidths(Table table, int rowIndex) =>
        TableGridNavigator.Cells(TableGridNavigator.Rows(table)[rowIndex].InnerRow!)
            .Select(c => c.InnerCell?.GetFirstChild<TableCellProperties>()?.GetFirstChild<TableCellWidth>()?.Width?.Value)
            .ToList();

    [Fact]
    public void SetColumnWidths_slices_a_skipped_cell_row_from_its_own_starting_grid_column()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.TableWithSkippedGridCells());
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                TableStructureEditor.SetColumnWidths(
                    doc, LayoutPart.Body, null, 0, new[] { 100, 200, 400, 800 });
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var table = BodyTable(reopened, 0);

            // Row 0 covers the whole grid, so it gets every width in order.
            Assert.Equal(new string?[] { "100", "200", "400", "800" }, CellWidths(table, 0));

            // Row 1 (gridAfter=2) covers only columns 0..1 - it must NOT be handed 400/800.
            Assert.Equal(new string?[] { "100", "200" }, CellWidths(table, 1));

            // Row 2 (gridBefore=1) STARTS at column 1: the offset is the whole point. Slicing from 0 would
            // have given it 100/200/400.
            Assert.Equal(new string?[] { "200", "400", "800" }, CellWidths(table, 2));

            Assert.Empty(OpenXmlErrors(reopened));
            Assert.Empty(TableGridConsistencyGuard.Find(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RemoveColumn_inside_a_trailing_skipped_run_shrinks_the_run_instead_of_a_cell()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.TableWithSkippedGridCells());
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                // Grid column 3 is a real cell in rows 0 and 2, but sits in row 1's skipped run.
                TableStructureEditor.RemoveColumn(doc, LayoutPart.Body, null, 0, 3);
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var table = BodyTable(reopened, 0);
            Assert.Equal(3, TableGridNavigator.GridColumnCount(table));

            Assert.Equal((0, 3, 0), Coverage(table, 0)); // lost a cell
            Assert.Equal((0, 2, 1), Coverage(table, 1)); // kept both cells, run 2 -> 1
            Assert.Equal((1, 2, 0), Coverage(table, 2)); // lost a cell

            Assert.Empty(OpenXmlErrors(reopened));
            Assert.Empty(TableGridConsistencyGuard.Find(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RemoveColumn_inside_a_leading_skipped_run_shrinks_the_run_instead_of_a_cell()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.TableWithSkippedGridCells());
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                // Grid column 0 is a real cell in rows 0 and 1, but is row 2's single skipped column.
                TableStructureEditor.RemoveColumn(doc, LayoutPart.Body, null, 0, 0);
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var table = BodyTable(reopened, 0);
            Assert.Equal(3, TableGridNavigator.GridColumnCount(table));

            Assert.Equal((0, 3, 0), Coverage(table, 0));
            Assert.Equal((0, 1, 2), Coverage(table, 1));
            Assert.Equal((0, 3, 0), Coverage(table, 2)); // run 1 -> 0, all three cells kept

            // A zero-valued gridBefore is removed rather than written out as w:val="0".
            var row2 = TableGridNavigator.Rows(table)[2].InnerRow!;
            Assert.Null(row2.GetFirstChild<TableRowProperties>()?.GetFirstChild<GridBefore>());

            Assert.Empty(OpenXmlErrors(reopened));
            Assert.Empty(TableGridConsistencyGuard.Find(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertColumn_append_grows_a_filler_rows_trailing_skipped_run_instead_of_adding_a_cell()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.TableWithSkippedGridCells());
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                // This table has no repeater and no w:tblHeader, so in plainText mode NO row is classified as
                // receiving content (see InsertColumn's per-row plan) - every row here is a filler, which is
                // exactly the case this test wants: the skipped-run rule applies unconditionally.
                TableStructureEditor.InsertColumn(
                    doc, LayoutPart.Body, null, 0, null,
                    new InsertColumnOptions { Mode = InsertColumnMode.PlainText, HeaderText = "New" });
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var table = BodyTable(reopened, 0);
            Assert.Equal(5, TableGridNavigator.GridColumnCount(table));

            // Row 0 covers the grid with no trailing run, so the new column lands at its end: a real cell.
            Assert.Equal((0, 5, 0), Coverage(table, 0));

            // Row 1's cells stop at grid column 2, so the new column falls inside its skipped run: the run
            // grows 2 -> 3 and the row looks exactly as it did, no cell added.
            Assert.Equal((0, 2, 3), Coverage(table, 1));

            // Row 2 has a LEADING run only and ends exactly at the grid's old end, so it too gains a cell.
            Assert.Equal((1, 4, 0), Coverage(table, 2));

            Assert.Empty(OpenXmlErrors(reopened));
            Assert.Empty(TableGridConsistencyGuard.Find(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertColumn_append_on_the_corpus_gridAfter_line_table_places_the_field_in_every_content_row()
    {
        // The headline gridAfter case (GitHub issue #9's history), against the stock layout that motivated
        // it. VatSpec's line-items table is
        // 11 grid columns: three rows of ten cells plus one skipped column, and one row of eleven cells.
        // Before this change every column operation refused the table outright.
        var path = CopyOfCorpus(Corpus.SalesInvoiceVatSpec);
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var before = TableGridNavigator.Coverage(
                    TableGridNavigator.Rows(BodyTable(doc, 2))[0].InnerRow!);
                Assert.Equal(1, before.After); // the fixture really is a gridAfter table

                var result = TableStructureEditor.InsertColumn(
                    doc, LayoutPart.Body, null, 2, null,
                    new InsertColumnOptions
                    {
                        Mode = InsertColumnMode.Field,
                        DataPath = "/Header/Line/ItemReferenceNo_Line",
                    });

                Assert.Equal(11, result.ColumnCountBefore);
                Assert.Equal(12, result.ColumnCountAfter);
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var table = BodyTable(reopened, 2);
            Assert.Equal(12, TableGridNavigator.GridColumnCount(table));

            // Every content row gained a real cell and KEPT its trailing filler column at the right edge;
            // no row silently grew its skipped run in place of the field that was asked for.
            Assert.Equal((0, 11, 1), Coverage(table, 0));
            Assert.Equal((0, 11, 1), Coverage(table, 1));
            Assert.Equal((0, 12, 0), Coverage(table, 3));

            Assert.Empty(OpenXmlErrors(reopened));
            Assert.Empty(TableGridConsistencyGuard.Find(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RemoveColumn_on_the_corpus_gridAfter_line_table_drops_the_skipped_column()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoiceVatSpec);
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                // Grid column 10 is the skipped column in rows 0..2 and part of row 3's last spanned cell.
                TableStructureEditor.RemoveColumn(doc, LayoutPart.Body, null, 2, 10);
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var table = BodyTable(reopened, 2);
            Assert.Equal(10, TableGridNavigator.GridColumnCount(table));
            Assert.Equal((0, 10, 0), Coverage(table, 0));
            Assert.Equal((0, 10, 0), Coverage(table, 1));
            Assert.Equal((0, 10, 0), Coverage(table, 3));

            Assert.Empty(OpenXmlErrors(reopened));
            Assert.Empty(TableGridConsistencyGuard.Find(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Guard_reports_no_violations_on_the_pristine_gridAfter_corpus_layouts()
    {
        // Non-tautology guard for the two tests above: their post-edit "no violations" assertions only mean
        // something if these layouts start clean. Down to one witness since QuantityExplosionofBOM left the
        // corpus (see Corpus.cs) — the remaining one is the harder 11-column case.
        foreach (var file in new[] { Corpus.SalesInvoiceVatSpec })
        {
            using var doc = WordprocessingDocument.Open(Corpus.Path(file), false);
            var violations = TableGridConsistencyGuard.Find(doc);
            Assert.True(violations.Count == 0,
                $"{file}: " + string.Join(" | ", violations.Select(v => v.Describe())));
        }
    }
}

/// <summary>
/// Tool-layer coverage for the table-structure tools via <see cref="TableTools"/> — the save-or-reject
/// envelope, the always-present <c>quickValidation</c>, and the file-untouched guarantee on rejection.
/// </summary>
public class TableStructureToolTests
{
    private const int LineTableIndex = 2;

    private static string CopyOfCorpus(string corpusFile)
    {
        var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"bcwl-tabletool-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    [Fact]
    public void Tool_set_column_widths_succeeds_with_quick_validation()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = TableTools.SetColumnWidths(
                path, LineTableIndex, "1000,2000,800,500,500,400,1300,700,1600");

            Assert.True(response.Ok, response.Error?.Message);
            var dto = Assert.IsType<TableEditResultDto>(response.Data);
            Assert.Equal("set_column_widths", dto.Operation);
            Assert.NotNull(dto.QuickValidation);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Tool_remove_column_succeeds()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = TableTools.RemoveColumn(path, LineTableIndex, column: 6);
            Assert.True(response.Ok, response.Error?.Message);
            var dto = Assert.IsType<TableEditResultDto>(response.Data);
            Assert.Equal(8, dto.ColumnCountAfter);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Tool_insert_column_interior_succeeds_and_reports_the_new_columns_grid_index()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = TableTools.InsertColumn(
                path, LineTableIndex, mode: "field", dataPath: "/Header/Line/TransHeaderAmount", atColumn: 2);

            Assert.True(response.Ok, response.Error?.Message);
            var dto = Assert.IsType<TableEditResultDto>(response.Data);
            Assert.Equal(2, dto.ColumnIndex);
            Assert.Equal(10, dto.ColumnCountAfter);
            Assert.True(dto.QuickValidation.Passed, "an interior insert must not introduce findings");

            using var reopened = WordprocessingDocument.Open(path, false);
            Assert.Empty(TableGridConsistencyGuard.Find(reopened));
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Tool_insert_column_interior_inside_a_spanned_content_cell_leaves_the_file_untouched()
    {
        var path = SyntheticLayout.Create(TableStructureEditorTests.SpannedHeaderTable);
        try
        {
            var before = File.ReadAllBytes(path);
            var response = TableTools.InsertColumn(path, tableIndex: 0, mode: "plainText", atColumn: 1);

            Assert.False(response.Ok);
            Assert.Equal("invalid_argument", response.Error!.Code);
            Assert.Contains("split_cells", response.Error.Hint, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(before, File.ReadAllBytes(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Tool_insert_column_append_field_succeeds()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = TableTools.InsertColumn(
                path, LineTableIndex, mode: "field", dataPath: "/Header/Line/TransHeaderAmount");

            Assert.True(response.Ok, response.Error?.Message);
            var dto = Assert.IsType<TableEditResultDto>(response.Data);
            Assert.Equal("insert_column", dto.Operation);
            Assert.Equal(10, dto.ColumnCountAfter);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- The grid-consistency guard's BEFORE/AFTER diff must not misfire on a
    // table that already has a pre-existing violation (see TableGridViolation's own remarks for the identity
    // rule this locks in) ----

    private static Table BodyTable(WordprocessingDocument doc, int index) =>
        doc.MainDocumentPart!.Document!.Body!.Descendants<Table>().ElementAt(index);

    /// <summary>A 3-grid-column table: row 0 is clean (3 cells); row 1 is pre-damaged, covering only 1 of the
    /// 3 declared columns (short by 2 - the magnitude that was ALWAYS rejected for
    /// remove_column, not just the one-short case that happened to work by accident).</summary>
    private const string TableWithPreExistingRaggedRow =
        "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"1000\"/><w:gridCol w:w=\"1000\"/><w:gridCol w:w=\"1000\"/></w:tblGrid>"
        + "<w:tr><w:tc><w:tcPr/><w:p/></w:tc><w:tc><w:tcPr/><w:p/></w:tc><w:tc><w:tcPr/><w:p/></w:tc></w:tr>"
        + "<w:tr><w:tc><w:tcPr/><w:p/></w:tc></w:tr>"
        + "</w:tbl>";

    [Fact]
    public void InsertColumn_succeeds_on_a_table_with_a_pre_existing_ragged_row_and_leaves_the_damage_untouched()
    {
        var path = SyntheticLayout.Create(TableWithPreExistingRaggedRow);
        try
        {
            // Confirm the fixture really is pre-damaged before the tool ever touches it.
            using (var pre = WordprocessingDocument.Open(path, false))
            {
                var preViolations = TableGridConsistencyGuard.Find(pre);
                Assert.Single(preViolations);
                Assert.Equal(TableGridViolationKind.CoverageMismatch, preViolations[0].Kind);
                Assert.Equal(1, preViolations[0].RowIndex);
            }

            // Regression: under the old description-string diff this was ALWAYS rejected (edit_would_break_table,
            // "likely a bug in the tool") because insert_column changes every row's expected grid count, making
            // the SAME pre-existing violation reword itself ("declares 3" -> "declares 4") and read as new.
            var response = TableTools.InsertColumn(path, tableIndex: 0, mode: "plainText");
            Assert.True(response.Ok, response.Error?.Message);

            using var reopened = WordprocessingDocument.Open(path, false);
            Assert.Equal(4, BodyTable(reopened, 0).GetFirstChild<TableGrid>()!.Elements<GridColumn>().Count());

            var violations = TableGridConsistencyGuard.Find(reopened);
            Assert.Single(violations); // the same pre-existing damage, still there, untouched by the edit
            Assert.Equal(TableGridViolationKind.CoverageMismatch, violations[0].Kind);
            Assert.Equal(1, violations[0].RowIndex);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RemoveColumn_succeeds_on_a_table_with_a_pre_existing_ragged_row_and_leaves_the_damage_untouched()
    {
        var path = SyntheticLayout.Create(TableWithPreExistingRaggedRow);
        try
        {
            // Regression: under the old description-string diff, a row short by TWO OR MORE columns (this
            // fixture) was ALWAYS rejected - unlike the one-short case, which happened to succeed only because
            // the shrunk grid coincidentally matched the malformed row's coverage exactly (the
            // "second-pass refinement"). remove_column skips malformed rows outright (TableStructureEditor.
            // RemoveColumn's "do not guess how to shrink it" branch), so the SAME violation persists - just
            // reworded ("declares 3" -> "declares 2") - and must stay allowed regardless of magnitude.
            var response = TableTools.RemoveColumn(path, tableIndex: 0, column: 0);
            Assert.True(response.Ok, response.Error?.Message);

            using var reopened = WordprocessingDocument.Open(path, false);
            Assert.Equal(2, BodyTable(reopened, 0).GetFirstChild<TableGrid>()!.Elements<GridColumn>().Count());

            var violations = TableGridConsistencyGuard.Find(reopened);
            Assert.Single(violations); // the same pre-existing damage, still there, untouched by the edit
            Assert.Equal(TableGridViolationKind.CoverageMismatch, violations[0].Kind);
            Assert.Equal(1, violations[0].RowIndex);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void GuardTableEdit_still_rejects_an_edit_that_introduces_new_raggedness_on_a_clean_table()
    {
        // STILL-GUARDS: the identity/count diff must keep catching genuinely NEW damage. Production table-
        // structure tools are correct-by-construction on documented inputs (see TableStructureEditor's own
        // class remarks) so none of them can be driven to actually desync a clean table through its public
        // tool surface; GuardTableEdit is internal precisely so a test can supply a deliberately corrupting
        // "mutate" callback and prove the real guard/diff/save-or-reject pipeline still rejects it, exactly
        // like the existing EditLockFor concurrency test uses the same internal-for-testing pattern.
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(rows: 2, cols: 3));
        try
        {
            var before = File.ReadAllBytes(path);

            var response = ToolGuards.GuardTableEdit(path, doc =>
            {
                var table = doc.MainDocumentPart!.Document!.Body!.Descendants<Table>().First();
                var row1 = table.Elements<TableRow>().ElementAt(1);
                row1.Elements<TableCell>().Last().Remove(); // desyncs row 1 from the 3-column grid

                return new TableEditResult
                {
                    Operation = "test_break_grid",
                    Part = "document.xml",
                    TableIndex = 0,
                    ColumnIndex = null,
                    RowsAffected = 1,
                    ColumnCountBefore = 3,
                    ColumnCountAfter = 3,
                    Summary = "test fixture: deliberately desyncs row 1",
                };
            });

            Assert.False(response.Ok);
            Assert.NotNull(response.Error);
            Assert.Equal("edit_would_break_table", response.Error!.Code);
            Assert.False(string.IsNullOrWhiteSpace(response.Error.Hint));
            Assert.Equal(before, File.ReadAllBytes(path)); // rejected before the atomic commit - file untouched
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RemoveControl_removing_a_row_above_a_pre_existing_ragged_row_does_not_falsely_reject_the_shift()
    {
        // ROW-SHIFT edge case: remove_control on a row-level control deletes that whole physical row (see
        // LayoutEditor.RemoveControl's block-level path - a row-level control's parent is the w:tbl, never a
        // TableCell, so it takes the block-level "wholesale removal" branch), shifting every later row's
        // RowIndex down by one. A pre-existing ragged row BELOW the one removed must not be misread as a NEW
        // violation just because it now reports a different RowIndex - it is the exact same damage, untouched.
        var body =
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"1000\"/><w:gridCol w:w=\"1000\"/></w:tblGrid>"
            + "<w:sdt><w:sdtPr><w:id w:val=\"777\"/></w:sdtPr><w:sdtContent>"
            + "<w:tr><w:tc><w:tcPr/><w:p/></w:tc><w:tc><w:tcPr/><w:p/></w:tc></w:tr>"
            + "</w:sdtContent></w:sdt>"
            + "<w:tr><w:tc><w:tcPr/><w:p/></w:tc></w:tr>" // ragged: covers 1 of 2 columns, at RowIndex 1
            + "</w:tbl>";
        var path = SyntheticLayout.Create(body);
        try
        {
            using (var pre = WordprocessingDocument.Open(path, false))
            {
                var preViolations = TableGridConsistencyGuard.Find(pre);
                Assert.Single(preViolations);
                Assert.Equal(1, preViolations[0].RowIndex);
            }

            var response = EditTools.RemoveControl(path, controlId: 777, keepText: false);
            Assert.True(response.Ok, response.Error?.Message);

            using var reopened = WordprocessingDocument.Open(path, false);
            var table = BodyTable(reopened, 0);
            Assert.Single(table.Elements<TableRow>()); // the row-level control's whole row is gone
            Assert.Empty(table.Elements<SdtRow>());

            var violations = TableGridConsistencyGuard.Find(reopened);
            Assert.Single(violations); // the same pre-existing damage, now at the shifted index - not "new"
            Assert.Equal(TableGridViolationKind.CoverageMismatch, violations[0].Kind);
            Assert.Equal(0, violations[0].RowIndex); // shifted down from 1 -> 0
        }
        finally
        {
            File.Delete(path);
        }
    }
}
