using BcWordLayout.Domain;
using BcWordLayout.McpHost;
using BcWordLayout.McpHost.Tools;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

/// <summary>
/// GitHub issue #76: an insert whose location resolves INSIDE a data-bound content control's own
/// <c>w:sdtContent</c> — <c>afterControl</c> on a cell-level field, <c>tableCell</c> on a cell that IS such a
/// control, an <c>atText</c> hit in its text — used to be written into that control and reported
/// <c>ok:true</c>. Word either refuses the file (field codes or a table inside a <c>&lt;w:text/&gt;</c>
/// control) or silently throws the content away on its next save (anything inside a <c>w15:dataBinding</c>
/// control). Every such location must now be REFUSED up front (ADR-0003), for every insert kind and both
/// binding shapes, with the file left untouched — while the legitimate neighbours (a plain cell, a cell in
/// a repeater ROW, an unbound cell-level wrapper) keep working exactly as before.
/// </summary>
public class BoundControlAnchorTests
{
    // StandardSalesQuote: /Header/QuoteValidToDate, a cell-level <w:dataBinding> + <w:text/> control.
    private const int QuoteValidToDateId = 1521969404;

    // StandardSalesQuote footer3.xml: /Header/CompanyEMail, a cell-level <w:dataBinding> + <w:text/> control.
    private const int FooterCompanyEMailId = 387000539;

    // InventoryOrderDetails: /Item/No_Item, a cell-level <w15:dataBinding> control with NO <w:text/>.
    private const int W15OnlyItemNoId = -1281949297;

    private static string CopyOfCorpus(string corpusFile)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-boundanchor-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    private static void AssertRefusedUntouched(ToolResponse response, string path, byte[] before)
    {
        Assert.False(response.Ok);
        Assert.Equal("invalid_argument", response.Error!.Code);
        Assert.Contains("data-bound", response.Error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("tableCell", response.Error.Hint, StringComparison.Ordinal);
        Assert.Contains("afterControl", response.Error.Hint, StringComparison.Ordinal);
        Assert.Equal(before, File.ReadAllBytes(path));
    }

    // ---- the exact reproductions from the issue, through the real tool surface ----

    [Fact]
    public void InsertPageNumber_afterControl_on_a_cell_level_plaintext_field_is_refused()
    {
        var path = CopyOfCorpus(Corpus.StandardSalesQuote);
        try
        {
            var before = File.ReadAllBytes(path);
            var response = EditTools.InsertPageNumber(path, "afterControl", controlId: QuoteValidToDateId);
            AssertRefusedUntouched(response, path, before);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertPageNumber_afterControl_on_a_cell_level_footer_field_is_refused()
    {
        var path = CopyOfCorpus(Corpus.StandardSalesQuote);
        try
        {
            var before = File.ReadAllBytes(path);
            var response = EditTools.InsertPageNumber(
                path, "afterControl", controlId: FooterCompanyEMailId, layoutPart: "footer", partName: "footer3.xml");
            AssertRefusedUntouched(response, path, before);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertField_afterControl_on_a_w15_only_cell_level_field_is_refused()
    {
        // The silent-drop variant: no <w:text/>, so neither OpenXmlValidator nor the plain-text guard has
        // anything to say - Word just discards the nested control on its next save.
        var path = CopyOfCorpus(Corpus.InventoryOrderDetails);
        try
        {
            var before = File.ReadAllBytes(path);
            var response = EditTools.InsertField(
                path, "/Item/Description_Item", "afterControl", controlId: W15OnlyItemNoId);
            AssertRefusedUntouched(response, path, before);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertPageNumber_and_InsertTable_into_a_tableCell_that_is_a_bound_control_are_refused()
    {
        var path = CopyOfCorpus(Corpus.StandardSalesQuote);
        try
        {
            // Verify the corpus shape the issue names rather than assume it: table 1, row 1, col 1 is a
            // cell-level control carrying a binding.
            using (var probe = WordprocessingDocument.Open(path, false))
            {
                var (cell, _) = LocationResolver.ResolveCellElement(
                    new Location { Type = LocationKind.TableCell, TableIndex = 1, Row = 1, Col = 1 }, probe);
                var sdtCell = Assert.IsType<SdtCell>(cell.Parent!.Parent);
                Assert.NotNull(sdtCell.SdtProperties!.GetFirstChild<DataBinding>());
            }

            var before = File.ReadAllBytes(path);
            AssertRefusedUntouched(EditTools.InsertPageNumber(path, "tableCell", tableIndex: 1, row: 1, col: 1), path, before);
            AssertRefusedUntouched(TableTools.InsertTable(path, 1, 1, "tableCell", tableIndex: 1, row: 1, col: 1), path, before);
            AssertRefusedUntouched(EditTools.InsertText(path, "x", "tableCell", tableIndex: 1, row: 1, col: 1), path, before);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Every_insert_kind_afterControl_on_a_cell_level_bound_field_is_refused()
    {
        var path = CopyOfCorpus(Corpus.StandardSalesQuote);
        try
        {
            var before = File.ReadAllBytes(path);
            AssertRefusedUntouched(
                EditTools.InsertField(path, "/Header/DocumentDate", "afterControl", controlId: QuoteValidToDateId), path, before);
            AssertRefusedUntouched(
                EditTools.InsertText(path, " ", "afterControl", controlId: QuoteValidToDateId), path, before);
            AssertRefusedUntouched(
                EditTools.InsertPicture(path, "/Header/CompanyPicture", "afterControl", controlId: QuoteValidToDateId), path, before);
            AssertRefusedUntouched(
                TableTools.InsertTable(path, 1, 1, "afterControl", controlId: QuoteValidToDateId), path, before);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- LocationResolver: both binding shapes, every location kind ----

    [Theory]
    [InlineData("<w:dataBinding w:xpath=\"/Header/CompanyName\" w:storeItemID=\"" + SyntheticLayout.GoodItemId + "\"/><w:text/>")]
    [InlineData("<w:dataBinding w:xpath=\"/Header/CompanyName\" w:storeItemID=\"" + SyntheticLayout.GoodItemId + "\"/>")]
    [InlineData("<w15:dataBinding w:xpath=\"/Header/CompanyName\" w:storeItemID=\"" + SyntheticLayout.GoodItemId + "\"/>")]
    public void Every_location_kind_resolving_inside_a_bound_cell_level_control_is_refused(string bindingMarkup)
    {
        var path = SyntheticLayout.Create(BoundCellLevelTable(42, bindingMarkup, "BOUND-CELL-TEXT"));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);

            foreach (var location in new[]
            {
                new Location { Type = LocationKind.AfterControl, ControlId = 42 },
                new Location { Type = LocationKind.TableCell, TableIndex = 0, Row = 0, Col = 0 },
                new Location { Type = LocationKind.AtText, SearchText = "BOUND-CELL-TEXT" },
            })
            {
                var ex = Assert.Throws<ArgumentException>(() => LocationResolver.Resolve(location, doc));
                Assert.Equal("location", ex.ParamName);
                Assert.Contains("id 42", ex.Message, StringComparison.Ordinal);
            }

            // The plain sibling cell in the same row is still a perfectly good target.
            var anchor = LocationResolver.Resolve(
                new Location { Type = LocationKind.TableCell, TableIndex = 0, Row = 0, Col = 1 }, doc);
            Assert.Equal("document.xml", anchor.PartName);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void AtText_matching_the_text_of_a_bound_block_level_control_is_refused()
    {
        var path = SyntheticLayout.Create(
            SyntheticLayout.BoundField("/Header/CompanyName", SyntheticLayout.GoodItemId).Replace(">x<", ">BLOCK-BOUND-TEXT<", StringComparison.Ordinal));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var ex = Assert.Throws<ArgumentException>(() => LocationResolver.Resolve(
                new Location { Type = LocationKind.AtText, SearchText = "BLOCK-BOUND-TEXT" }, doc));
            Assert.Equal("location", ex.ParamName);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void AtText_matching_the_text_of_an_inline_bound_control_still_resolves_outside_it()
    {
        // An inline (run-level) control's paragraph is OUTSIDE the control - the anchor appends to that
        // paragraph after the control, which is a legitimate position, not the control's own content.
        var path = SyntheticLayout.Create(
            "<w:p><w:sdt><w:sdtPr><w:id w:val=\"7\"/>"
            + $"<w:dataBinding w:xpath=\"/Header/CompanyName\" w:storeItemID=\"{SyntheticLayout.GoodItemId}\"/><w:text/>"
            + "</w:sdtPr><w:sdtContent><w:r><w:t>INLINE-BOUND-TEXT</w:t></w:r></w:sdtContent></w:sdt></w:p>");
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var anchor = LocationResolver.Resolve(new Location { Type = LocationKind.AtText, SearchText = "INLINE-BOUND-TEXT" }, doc);
            var marker = new Run(new Text("MARKER"));
            anchor.InsertInline(marker);
            Assert.Empty(marker.Ancestors<SdtElement>());
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- what must keep working: containers are not bound fields ----

    [Fact]
    public void TableCell_inside_a_repeater_row_still_resolves()
    {
        // The repeatingSection/repeatingSectionItem pair carries a w15:dataBinding of its own, but it is a
        // CONTAINER: filling a cell of the row template is the bread-and-butter repeater edit.
        var path = SyntheticLayout.Create(
            SyntheticLayout.RepeaterTable("/Header/Line", "/Header/Line/ItemNo", SyntheticLayout.GoodItemId),
            datasetXml: SyntheticLayout.DatasetWithLineXml);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var anchor = LocationResolver.Resolve(
                new Location { Type = LocationKind.TableCell, TableIndex = 0, Row = 1, Col = 0 }, doc);
            var marker = new Run(new Text("MARKER-IN-REPEATER-ROW"));
            anchor.InsertInline(marker);

            Assert.NotNull(marker.Ancestors<SdtRow>().FirstOrDefault());
            Assert.DoesNotContain(marker.Ancestors<SdtElement>(), s => s.SdtProperties?.GetFirstChild<DataBinding>() is not null);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertField_tableCell_in_a_real_line_items_repeater_row_still_succeeds()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            int tableIndex, row, col;
            using (var probe = WordprocessingDocument.Open(path, false))
            {
                // A bare w:tc of a row-level repeater item's row (its row sits in a repeater's sdtContent).
                var body = probe.MainDocumentPart!.Document!.Body!;
                var target = body.Descendants<TableCell>().First(c => c.Parent is TableRow { Parent: SdtContentRow });
                var table = target.Ancestors<Table>().First();
                tableIndex = body.Descendants<Table>().ToList().IndexOf(table);
                row = TableGridNavigator.Rows(table).Select((s, i) => (s, i))
                    .First(x => ReferenceEquals(x.s.InnerRow, target.Parent)).i;
                col = TableGridNavigator.Cells((TableRow)target.Parent!).Select((s, i) => (s, i))
                    .First(x => ReferenceEquals(x.s.InnerCell, target)).i;
            }

            var response = EditTools.InsertText(path, "IN-REPEATER-ROW", "tableCell", tableIndex: tableIndex, row: row, col: col);
            Assert.True(response.Ok, response.Error?.Message);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void AfterControl_on_an_unbound_cell_level_wrapper_still_anchors_inside_its_cell()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.TableWithCellLevelControl(id: 99, text: "unbound"));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var anchor = LocationResolver.Resolve(new Location { Type = LocationKind.AfterControl, ControlId = 99 }, doc);
            var marker = new Run(new Text("MARKER"));
            anchor.InsertInline(marker);
            Assert.IsType<SdtContentCell>(marker.Ancestors<TableCell>().First().Parent);
        }
        finally
        {
            File.Delete(path);
        }
    }

    /// <summary>
    /// One row: a cell-level control carrying <paramref name="bindingMarkup"/> in its <c>w:sdtPr</c> (any of
    /// the three binding shapes real layouts use), then one plain sibling cell.
    /// </summary>
    private static string BoundCellLevelTable(int id, string bindingMarkup, string text) =>
        "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2000\"/><w:gridCol w:w=\"2000\"/></w:tblGrid>"
        + "<w:tr>"
        + $"<w:sdt><w:sdtPr><w:id w:val=\"{id}\"/>{bindingMarkup}</w:sdtPr>"
        + $"<w:sdtContent><w:tc><w:tcPr/><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc></w:sdtContent></w:sdt>"
        + "<w:tc><w:tcPr/><w:p><w:r><w:t>plain-cell</w:t></w:r></w:p></w:tc>"
        + "</w:tr></w:tbl>";
}
