using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;
using BcWordLayout.McpHost.Tools;
using BcWordLayout.Merge;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

/// <summary>
/// Which rows of a table <c>insert_column</c> puts the header cell and the bound field in (GitHub issue
/// #82), against copies of the stock layouts the issue was reported on: a field belongs only in the rows
/// of the repeater whose data item contains it (the 1304/1306 lines tables also repeat
/// <c>/Header/ReportTotalsLine</c>), a plain table's header rows are its caption rows, and header/data
/// arguments with nowhere to go are refused rather than dropped.
/// </summary>
public class InsertColumnRowRoleTests
{
    private const int QuoteInfoTable = 1;  // caption-over-value: 5 label cells over 5 field cells
    private const int QuoteLinesTable = 2; // tblHeader row, Line repeater, ReportTotalsLine repeater, totals

    private static string CopyOfCorpus(string corpusFile)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-colroles-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    private static Table BodyTable(WordprocessingDocument doc, int index) =>
        doc.MainDocumentPart!.Document!.Body!.Descendants<Table>().ElementAt(index);

    private static IReadOnlyList<TableGridNavigator.CellSlot> RowCells(Table table, int row) =>
        TableGridNavigator.Cells(TableGridNavigator.Rows(table)[row].InnerRow!);

    /// <summary>Leaf names of every bound control in <paramref name="table"/> whose binding ends in <paramref name="leaf"/>.</summary>
    private static int CountBound(OpenXmlElement scope, string leaf) =>
        scope.Descendants<SdtElement>().Count(s => BindingXPath.LeafName(SdtInspector.ReadXPath(s)) == leaf);

    private static int XPathFallbacks(string path) =>
        FullValidator.Full(path).Findings.Count(f => f.Check == "xpath-fallback");

    [Fact]
    public void Demo_insert_column_on_the_quote_lines_table_binds_only_the_Line_row_and_adds_no_xpath_fallback()
    {
        // The talk demo: an Item Reference column at grid 1 of the stock quote's lines table. Before #82 the
        // field also landed inside the ReportTotalsLine repeater, where it renders Line 1's value beside
        // every totals line and full validation reports a new xpath-fallback warning for it (see the next
        // test, which proves this comparison can fail).
        var path = CopyOfCorpus(Corpus.StandardSalesQuote);
        try
        {
            var fallbacksBefore = XPathFallbacks(path);

            var response = TableTools.InsertColumn(
                path, QuoteLinesTable, mode: "field", dataPath: "/Header/Line/ItemReferenceNo_Line", atColumn: 1);
            Assert.True(response.Ok, response.Error?.Message);

            using (var doc = WordprocessingDocument.Open(path, false))
            {
                var table = BodyTable(doc, QuoteLinesTable);
                Assert.Equal(1, CountBound(doc.MainDocumentPart!.Document!, "ItemReferenceNo_Line"));
                Assert.Contains(
                    TableStructureEditorTests.BoundXPathsInRepeater(table, "Line"),
                    x => BindingXPath.LeafName(x) == "ItemReferenceNo_Line");
                Assert.DoesNotContain(
                    TableStructureEditorTests.BoundXPathsInRepeater(table, "ReportTotalsLine"),
                    x => BindingXPath.LeafName(x) == "ItemReferenceNo_Line");

                // Row 0 (w:tblHeader) gets the header text; the totals repeater's row an empty cell.
                Assert.False(string.IsNullOrWhiteSpace(RowCells(table, 0)[1].CellChild.InnerText));
                Assert.Equal(string.Empty, RowCells(table, 4)[1].CellChild.InnerText);

                Assert.Empty(TableGridConsistencyGuard.Find(doc));
                Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc));
            }

            Assert.Equal(fallbacksBefore, XPathFallbacks(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void The_pre_82_placement_of_the_demo_field_in_the_totals_repeater_row_is_reported_by_full_validation()
    {
        // What the comparison above guards against: the same Item Reference field placed in the
        // ReportTotalsLine row, as insert_column did before #82, prints Line 1's value beside every totals
        // line. Simulated by rebinding that row's description control; full validation must report it once,
        // counting every totals row it renders in (GitHub issue #104 kept this signal).
        const int DescriptionReportTotalsLineId = 1468313313;
        const string LineField = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line[1]/ns0:ItemReferenceNo_Line[1]";
        var path = CopyOfCorpus(Corpus.StandardSalesQuote);
        try
        {
            var fallbacksBefore = XPathFallbacks(path);
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var control = Assert.Single(
                    doc.MainDocumentPart!.Document!.Body!.Descendants<SdtElement>(),
                    s => SdtInspector.ReadControlId(s) == DescriptionReportTotalsLineId);
                SdtInspector.FindBinding(control)!.SetAttribute(new OpenXmlAttribute("w", "xpath", OoxmlNames.W, LineField));
            }

            var findings = FullValidator.Full(path).Findings.Where(f => f.Check == "xpath-fallback").ToList();

            Assert.Equal(fallbacksBefore + 1, findings.Count);
            var finding = Assert.Single(findings, f => f.Location == $"document.xml: {LineField}");
            Assert.Equal(FindingSeverity.Warning, finding.Severity);
            Assert.StartsWith(
                "Field '/Header/Line/ItemReferenceNo_Line' lies outside every '/Header/Line' repeater, so it shows "
                + "the first of 3 'Line' records (3 evaluation(s),",
                finding.Message,
                StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Appending_a_Line_field_keeps_the_quotes_totals_repeater_row_right_anchored()
    {
        // The issue's first repro. The ReportTotalsLine row is a filler now, so the totals-block rule (leading
        // empty cells, content on the right edge) widens its amount cell instead of giving it a new cell.
        var path = CopyOfCorpus(Corpus.StandardSalesQuote);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var totalsCellsBefore = RowCells(BodyTable(doc, QuoteLinesTable), 4).Count;
            var amountSpanBefore = RowCells(BodyTable(doc, QuoteLinesTable), 4)[^1].GridSpan;

            TableStructureEditor.InsertColumn(
                doc, LayoutPart.Body, null, QuoteLinesTable, atColumn: null,
                new InsertColumnOptions
                {
                    Mode = InsertColumnMode.Field,
                    DataPath = "/Header/Line/AmountIncludingVAT_Line",
                    HeaderLabelPath = "/Header/Line/AmountIncludingVAT_Line_Lbl",
                });

            var table = BodyTable(doc, QuoteLinesTable);
            Assert.DoesNotContain(
                TableStructureEditorTests.BoundXPathsInRepeater(table, "ReportTotalsLine"),
                x => BindingXPath.LeafName(x) == "AmountIncludingVAT_Line");
            Assert.Equal(1, CountBound(table, "AmountIncludingVAT_Line"));
            Assert.Equal(1, CountBound(RowCells(table, 0)[^1].CellChild, "AmountIncludingVAT_Line_Lbl"));

            var totalsRow = RowCells(table, 4);
            Assert.Equal(totalsCellsBefore, totalsRow.Count);
            Assert.Equal(amountSpanBefore + 1, totalsRow[^1].GridSpan);
            Assert.Empty(TableGridConsistencyGuard.Find(doc));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Appending_a_Line_field_to_the_VatSpec_lines_table_leaves_its_totals_repeater_row_without_it()
    {
        // The issue's second repro (Shipment Date beside every totals line). That row has a trailing
        // w:gridAfter run, so as a filler it absorbs the new column without gaining a cell at all.
        var path = CopyOfCorpus(Corpus.SalesInvoiceVatSpec);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var totalsCellsBefore = RowCells(BodyTable(doc, 2), 2).Count;

            TableStructureEditor.InsertColumn(
                doc, LayoutPart.Body, null, 2, atColumn: null,
                new InsertColumnOptions { Mode = InsertColumnMode.Field, DataPath = "/Header/Line/ShipmentDate_Line" });

            var table = BodyTable(doc, 2);
            Assert.Equal(1, CountBound(table, "ShipmentDate_Line"));
            Assert.DoesNotContain(
                TableStructureEditorTests.BoundXPathsInRepeater(table, "ReportTotalsLine"),
                x => BindingXPath.LeafName(x) == "ShipmentDate_Line");
            Assert.Equal(totalsCellsBefore, RowCells(table, 2).Count);
            Assert.Empty(TableGridConsistencyGuard.Find(doc));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_field_outside_every_repeater_of_a_two_repeater_table_is_refused_naming_the_repeaters()
    {
        // /Header/YourReference is a parent-item field: in a table repeating both Line and ReportTotalsLine
        // there is no telling which row it is meant for.
        var path = CopyOfCorpus(Corpus.StandardSalesQuote);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var ex = Assert.Throws<ArgumentException>(() => TableStructureEditor.InsertColumn(
                doc, LayoutPart.Body, null, QuoteLinesTable, atColumn: null,
                new InsertColumnOptions { Mode = InsertColumnMode.Field, DataPath = "/Header/YourReference" }));

            Assert.Equal("dataPath", ex.ParamName);
            Assert.Contains("'/Header/Line'", ex.Message, StringComparison.Ordinal);
            Assert.Contains("'/Header/ReportTotalsLine'", ex.Message, StringComparison.Ordinal);
            Assert.Equal(8, TableGridNavigator.GridColumnCount(BodyTable(doc, QuoteLinesTable))); // untouched
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_parent_items_field_in_a_single_repeater_table_goes_into_that_repeaters_row()
    {
        // The tool contract allows a parent/Header field in a lines column. With one repeater the target row
        // is unambiguous: the purchase order's lines table repeats only Purchase_Line.
        var path = CopyOfCorpus(Corpus.StandardPurchaseOrder);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            TableStructureEditor.InsertColumn(
                doc, LayoutPart.Body, null, 3, atColumn: null,
                new InsertColumnOptions { Mode = InsertColumnMode.Field, DataPath = "/Purchase_Header/BuyFromAddr1" });

            var table = BodyTable(doc, 3);
            Assert.Contains(
                TableStructureEditorTests.BoundXPathsInRepeater(table, "Purchase_Line"),
                x => BindingXPath.LeafName(x) == "BuyFromAddr1");
            Assert.Equal(1, CountBound(table, "BuyFromAddr1"));
            Assert.Empty(TableGridConsistencyGuard.Find(doc));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void PlainText_header_text_on_a_caption_over_value_table_lands_in_the_caption_row()
    {
        // The issue's third repro: "EXTRACOL" used to appear nowhere in document.xml.
        var path = CopyOfCorpus(Corpus.StandardSalesQuote);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            TableStructureEditor.InsertColumn(
                doc, LayoutPart.Body, null, QuoteInfoTable, atColumn: null,
                new InsertColumnOptions { Mode = InsertColumnMode.PlainText, HeaderText = "EXTRACOL" });

            var table = BodyTable(doc, QuoteInfoTable);
            Assert.Equal("EXTRACOL", RowCells(table, 0)[^1].CellChild.InnerText);
            Assert.Equal(string.Empty, RowCells(table, 1)[^1].CellChild.InnerText);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Field_with_a_header_label_on_a_caption_over_value_table_binds_label_over_field()
    {
        // Before #82 the field was bound into BOTH rows and no label was emitted. (/Header/YourReference__Lbl,
        // double underscore included, is the stock quote schema's real column name - not a misspelling.)
        var path = CopyOfCorpus(Corpus.StandardSalesQuote);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var body = doc.MainDocumentPart!.Document!.Body!;
            var fieldsBefore = CountBound(body, "YourReference");

            TableStructureEditor.InsertColumn(
                doc, LayoutPart.Body, null, QuoteInfoTable, atColumn: null,
                new InsertColumnOptions
                {
                    Mode = InsertColumnMode.Field,
                    DataPath = "/Header/YourReference",
                    HeaderLabelPath = "/Header/YourReference__Lbl",
                });

            var table = BodyTable(doc, QuoteInfoTable);
            Assert.Equal(1, CountBound(RowCells(table, 0)[^1].CellChild, "YourReference__Lbl"));
            Assert.Equal(1, CountBound(RowCells(table, 1)[^1].CellChild, "YourReference"));
            Assert.Equal(fieldsBefore + 1, CountBound(body, "YourReference"));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_header_label_path_that_does_not_resolve_is_refused_even_in_plainText_mode()
    {
        var path = CopyOfCorpus(Corpus.StandardSalesQuote);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var ex = Assert.Throws<ArgumentException>(() => TableStructureEditor.InsertColumn(
                doc, LayoutPart.Body, null, QuoteInfoTable, atColumn: null,
                new InsertColumnOptions { Mode = InsertColumnMode.PlainText, HeaderLabelPath = "/Header/NoSuchCaption_Lbl" }));

            Assert.Equal("headerLabelPath", ex.ParamName);
            Assert.Contains("NoSuchCaption_Lbl", ex.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Tool_header_text_on_a_table_with_no_header_row_is_refused_and_the_file_untouched()
    {
        // The quote's address block (table 0) is all field rows: no tblHeader, no repeater, no caption row,
        // so a headerText has nowhere to go. It used to return ok:true and drop it.
        var path = CopyOfCorpus(Corpus.StandardSalesQuote);
        try
        {
            var before = File.ReadAllBytes(path);
            var response = TableTools.InsertColumn(path, tableIndex: 0, mode: "plainText", headerText: "Notes");

            Assert.False(response.Ok);
            Assert.Equal("invalid_argument", response.Error!.Code);
            Assert.Contains("no header row", response.Error.Message, StringComparison.Ordinal);
            Assert.Contains("caption rows", response.Error.Hint, StringComparison.Ordinal);
            Assert.Equal(before, File.ReadAllBytes(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Header_text_on_a_literal_text_table_is_refused()
    {
        // Literal text is not a caption: a hand-typed row is indistinguishable from a row of values.
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(rows: 3, cols: 3));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var ex = Assert.Throws<ArgumentException>(() => TableStructureEditor.InsertColumn(
                doc, LayoutPart.Body, null, 0, atColumn: null,
                new InsertColumnOptions { Mode = InsertColumnMode.PlainText, HeaderText = "Notes" }));
            Assert.Equal("headerText", ex.ParamName);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
