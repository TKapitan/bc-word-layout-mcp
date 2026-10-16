using BcWordLayout.Domain;
using BcWordLayout.McpHost.Tools;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

/// <summary>
/// Covers <see cref="PlainTextNestingGuard"/> — the detector for content Word forbids inside a plain-text
/// content control (a nested control, a field code, a table, a drawing): combinations Word rejects as corrupt
/// but OpenXmlValidator accepts. Every clean corpus layout must report zero; a layout with any of those inside
/// a plain-text control must report it.
/// </summary>
public class PlainTextNestingGuardTests
{
    private const int PlainTextControlId = 77;

    [Theory]
    [InlineData(Corpus.SalesInvoice)]
    [InlineData(Corpus.InventoryOrderDetails)]
    [InlineData(Corpus.StandardStatement)]
    [InlineData(Corpus.StandardSalesQuote)]
    [InlineData(Corpus.StandardPurchaseOrder)]
    [InlineData(Corpus.SalespersonCommission)]
    [InlineData(Corpus.SalesInvoiceVatSpec)]
    [InlineData(Corpus.JobQuote)]
    [InlineData(Corpus.PaymentPracticeByPeriod)]
    [InlineData(Corpus.SubcontractorDispatchList)]
    public void Find_reports_nothing_for_a_clean_corpus_layout(string corpusFile)
    {
        using var doc = WordprocessingDocument.Open(Corpus.Path(corpusFile), false);
        Assert.Empty(PlainTextNestingGuard.Find(doc));
    }

    [Fact]
    public void Find_reports_a_content_control_nested_inside_a_real_cell_level_plaintext_control()
    {
        // YourReference_Lbl (id -1130623254) is a cell-level plain-text control. The insert tools refuse to
        // anchor inside it (GitHub issue #76), so the nesting is built directly here - the guard is the
        // backstop for whatever route produces it.
        const int cellLevelPlainTextLabelId = -1130623254;
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-ptguard-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(Corpus.SalesInvoice), path, overwrite: true);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var label = doc.MainDocumentPart!.Document!.Body!.Descendants<SdtCell>()
                .Single(s => s.SdtProperties?.GetFirstChild<SdtId>()?.Val?.Value == cellLevelPlainTextLabelId);
            label.Descendants<Paragraph>().Last().AppendChild(
                new SdtRun(new SdtProperties(new SdtId { Val = 4242 }), new SdtContentRun(new Run(new Text("nested")))));

            var offender = Assert.Single(PlainTextNestingGuard.Find(doc));
            Assert.Equal(cellLevelPlainTextLabelId, offender.OuterId);
            Assert.Equal(4242, offender.InnerId);
            Assert.Null(offender.InnerElement);
            Assert.Equal("document.xml", offender.Part);
            Assert.Contains("nested inside plain-text", offender.Describe(), StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(path);
        }
    }

    public static TheoryData<string, string> NonRunContent => new()
    {
        { "fldChar", "w:fldChar" },
        { "fldSimple", "w:fldSimple" },
        { "tbl", "w:tbl" },
        { "drawing", "w:drawing" },
    };

    [Theory]
    [MemberData(nameof(NonRunContent))]
    public void Find_reports_non_run_content_inside_a_plaintext_control(string kind, string expectedName)
    {
        var path = SyntheticLayout.Create(PlainTextCellLevelTable(bound: true));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var cell = doc.MainDocumentPart!.Document!.Body!.Descendants<SdtCell>().Single().SdtContentCell!
                .GetFirstChild<TableCell>()!;
            var paragraph = cell.Elements<Paragraph>().Last();

            switch (kind)
            {
                case "fldChar":
                    paragraph.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }));
                    paragraph.AppendChild(new Run(new FieldCode(" PAGE ")));
                    paragraph.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
                    break;
                case "fldSimple":
                    paragraph.AppendChild(new SimpleField(new Run(new Text("1"))) { Instruction = " PAGE " });
                    break;
                case "tbl":
                    cell.InsertBefore(
                        new Table(new TableProperties(), new TableGrid(new GridColumn { Width = "1000" }),
                            new TableRow(new TableCell(new Paragraph()))),
                        paragraph);
                    break;
                case "drawing":
                    paragraph.AppendChild(new Run(new Drawing()));
                    break;
            }

            var offenders = PlainTextNestingGuard.Find(doc);
            Assert.NotEmpty(offenders);
            Assert.All(offenders, o => Assert.Equal(PlainTextControlId, o.OuterId));
            Assert.Contains(offenders, o => o.InnerElement is not null && o.InnerElement.Contains(expectedName, StringComparison.Ordinal));
            Assert.Contains(offenders, o => o.Describe().Contains(expectedName, StringComparison.Ordinal));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Find_ignores_the_same_content_outside_any_plaintext_control()
    {
        var path = SyntheticLayout.Create(
            "<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r>"
            + "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>"
            + SyntheticLayout.SimpleTable(rows: 1, cols: 1));
        try
        {
            using var doc = WordprocessingDocument.Open(path, false);
            Assert.Empty(PlainTextNestingGuard.Find(doc));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertPageNumber_into_an_unbound_plaintext_cell_control_is_rejected_by_the_backstop()
    {
        // An UNBOUND w:text control is not a data-bound field, so LocationResolver lets tableCell anchor
        // inside it - but the plain-text rule still applies, and the field runs would make Word declare the
        // file corrupt. The mutation gate must catch it and leave the file byte-for-byte untouched.
        var path = SyntheticLayout.Create(PlainTextCellLevelTable(bound: false));
        try
        {
            var before = File.ReadAllBytes(path);

            var response = EditTools.InsertPageNumber(path, "tableCell", tableIndex: 0, row: 0, col: 0);

            Assert.False(response.Ok);
            Assert.Equal("edit_would_corrupt", response.Error!.Code);
            Assert.Contains("w:fldChar", response.Error.Hint, StringComparison.Ordinal);
            Assert.Equal(before, File.ReadAllBytes(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    /// <summary>
    /// One row: a cell-level plain-text (<c>&lt;w:text/&gt;</c>) control — bound to <c>/Header/CompanyName</c>
    /// or not — wrapping its <c>w:tc</c>, then one plain sibling cell.
    /// </summary>
    private static string PlainTextCellLevelTable(bool bound) =>
        "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2000\"/><w:gridCol w:w=\"2000\"/></w:tblGrid>"
        + "<w:tr>"
        + $"<w:sdt><w:sdtPr><w:id w:val=\"{PlainTextControlId}\"/>"
        + (bound
            ? $"<w:dataBinding w:xpath=\"/Header/CompanyName\" w:storeItemID=\"{SyntheticLayout.GoodItemId}\"/>"
            : string.Empty)
        + "<w:text/></w:sdtPr>"
        + "<w:sdtContent><w:tc><w:tcPr/><w:p><w:r><w:t>plain-text</w:t></w:r></w:p></w:tc></w:sdtContent></w:sdt>"
        + "<w:tc><w:tcPr/><w:p><w:r><w:t>plain-cell</w:t></w:r></w:p></w:tc>"
        + "</w:tr></w:tbl>";
}
