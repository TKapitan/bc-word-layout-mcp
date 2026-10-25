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
/// The three row tools (<c>insert_repeater_row</c>, <c>insert_subtotal_row</c>, <c>insert_table_row</c>) give each
/// new cell the run formatting of the cell in the same grid column of the row beside it, as Word does for an
/// inserted row, with <c>format</c> on top (<see cref="InheritedRunFormat.ForRowCells"/>). Sandbox pack P04 (BC 28)
/// printed a detail row the tools built bare at 11 pt under Job Quote's Segoe UI 8 pt task rows.
/// </summary>
[Collection("label-convention-seam")]
public class RowFormatInheritanceTests
{
    private const string PlanningLine = "/Job/Job_Task/Job_Planning_Line";

    private static string CopyOfCorpus(string corpusFile)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-rowfmt-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    private static int RepeaterId(WordprocessingDocument doc, string dataItem) =>
        LayoutReader.Read(doc).Controls
            .Single(c => c.Kind == ControlKind.Repeater && c.Alias == $"#Nav: {dataItem}").SdtId!.Value;

    private static SdtRow Control(WordprocessingDocument doc, int id) =>
        doc.MainDocumentPart!.Document!.Body!.Descendants<SdtRow>().Single(s => SdtInspector.ReadControlId(s) == id);

    /// <summary>The cell-by-cell spec of a Job Planning Line detail row covering <paramref name="gridCount"/> columns.</summary>
    private static List<RepeaterRowCell> PlanningLineCells(int gridCount) =>
    [
        new() { Span = gridCount - 3 },
        new() { Fields = [$"{PlanningLine}/Number"] },
        new() { Fields = [$"{PlanningLine}/Quantity"], Alignment = "right" },
        new() { Fields = [$"{PlanningLine}/TotalPrice"], Alignment = "right" },
    ];

    /// <summary>
    /// JobQuote's Job_Task repeater id and its table's grid column count. The task item holds no row of its own,
    /// only the nested Job_Planning_Line repeater, so a row appended to it follows the planning line's row.
    /// </summary>
    private static (int TaskRepeaterId, int GridCount) TaskRepeater(WordprocessingDocument doc)
    {
        var taskId = RepeaterId(doc, "/Job/Job_Task");
        var table = Control(doc, taskId).Ancestors<Table>().First();
        return (taskId, TableGridNavigator.GridColumnCount(table));
    }

    private static string? Size(OpenXmlElement? rPr) => rPr?.GetFirstChild<FontSize>()?.Val?.Value;

    private static string? Font(OpenXmlElement? rPr) => rPr?.GetFirstChild<RunFonts>()?.Ascii?.Value;

    private static bool IsBold(OpenXmlElement? rPr) =>
        rPr?.GetFirstChild<Bold>() is { } b && (b.Val is null || b.Val.Value);

    [Fact]
    public void A_detail_row_after_Job_Quotes_planning_line_row_takes_its_Segoe_UI_8_pt()
    {
        var path = CopyOfCorpus(Corpus.JobQuote);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var (taskId, gridCount) = TaskRepeater(doc);

            var result = LayoutEditor.InsertRepeaterRow(doc, taskId, PlanningLine, PlanningLineCells(gridCount));

            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc));
            Assert.Contains("takes the formatting of the cell above it (Segoe UI, 8 pt)", result.Summary, StringComparison.Ordinal);

            var row = Control(doc, result.ControlId).Descendants<TableRow>().Single();
            var controls = row.Descendants<SdtRun>().ToList();
            Assert.Equal(3, controls.Count);
            foreach (var control in controls)
            {
                var own = control.SdtProperties!.GetFirstChild<RunProperties>();
                Assert.Equal(("Segoe UI", "16"), (Font(own), Size(own)));
                Assert.All(control.Descendants<Run>(), run => Assert.Equal("16", Size(run.RunProperties)));
            }

            // Every paragraph mark too: it is what sizes the empty indent cell's line.
            Assert.All(row.Descendants<Paragraph>(), paragraph =>
                Assert.Equal("16", Size(paragraph.ParagraphProperties?.ParagraphMarkRunProperties)));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Format_overrides_what_the_row_takes_member_by_member()
    {
        var path = CopyOfCorpus(Corpus.JobQuote);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var (taskId, gridCount) = TaskRepeater(doc);

            var result = LayoutEditor.InsertRepeaterRow(
                doc, taskId, PlanningLine, PlanningLineCells(gridCount), new TextFormat { Bold = true, FontSizePoints = 9 });

            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc));
            var row = Control(doc, result.ControlId).Descendants<TableRow>().Single();
            foreach (var control in row.Descendants<SdtRun>())
            {
                var own = control.SdtProperties!.GetFirstChild<RunProperties>();
                Assert.Equal(("Segoe UI", "18"), (Font(own), Size(own)));
                Assert.True(IsBold(own));
            }

            Assert.All(row.Descendants<Paragraph>(), paragraph =>
            {
                var mark = paragraph.ParagraphProperties?.ParagraphMarkRunProperties;
                Assert.Equal(("Segoe UI", "18"), (Font(mark), Size(mark)));
                Assert.True(IsBold(mark));
            });
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_repeater_item_with_no_row_of_its_own_gives_nothing()
    {
        // Sandbox item P04's shape: Job Quote's own planning-line repeater removed, so the task item is empty. The
        // row above the repeater is the bold, underlined caption row, which is no model for a detail row.
        var path = CopyOfCorpus(Corpus.JobQuote);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            LayoutEditor.RemoveControl(doc, RepeaterId(doc, PlanningLine), keepText: false);
            var (taskId, gridCount) = TaskRepeater(doc);

            var result = LayoutEditor.InsertRepeaterRow(doc, taskId, PlanningLine, PlanningLineCells(gridCount));

            Assert.DoesNotContain("takes the formatting", result.Summary, StringComparison.Ordinal);
            var row = Control(doc, result.ControlId).Descendants<TableRow>().Single();
            Assert.All(row.Descendants<SdtRun>(), control =>
                Assert.Null(control.SdtProperties!.GetFirstChild<RunProperties>()));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_totals_row_after_Job_Quotes_lines_takes_the_last_line_rows_formatting()
    {
        var path = CopyOfCorpus(Corpus.JobQuote);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var taskRepeater = Control(doc, RepeaterId(doc, "/Job/Job_Task"));
            var table = taskRepeater.Ancestors<Table>().First();
            var tableIndex = TableGridNavigator.Tables(doc.MainDocumentPart!.Document!.Body!).ToList().IndexOf(table);
            var gridCount = TableGridNavigator.GridColumnCount(table);

            // Rows: the caption row, the planning line's row (inside the task repeater), the job total row. The new
            // row goes between the last two, right after the lines.
            Assert.Equal(3, TableGridNavigator.Rows(table).Count);
            var result = TableStructureEditor.InsertStaticRow(
                doc, LayoutPart.Body, null, tableIndex, atRow: 2,
                [new RepeaterRowCell { Span = gridCount - 1 }, new RepeaterRowCell { Fields = ["/Job/No_Job"], Alignment = "right" }]);

            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc));
            Assert.Contains("the cell above it (Segoe UI, 8 pt)", result.Summary, StringComparison.Ordinal);
            var row = TableGridNavigator.Rows(table)[2].InnerRow!;
            var own = row.Descendants<SdtRun>().Single().SdtProperties!.GetFirstChild<RunProperties>();
            Assert.Equal(("Segoe UI", "16"), (Font(own), Size(own)));
            Assert.False(IsBold(own));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_spacer_row_in_Salesperson_Commission_takes_the_7_pt_its_line_height_comes_from()
    {
        var path = CopyOfCorpus(Corpus.SalespersonCommission);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var parentId = LayoutReader.Read(doc).Controls
                .Single(c => c.Kind == ControlKind.Repeater && c.ParentRepeater is null).SdtId!.Value;

            var result = LayoutEditor.InsertSubtotalRow(doc, parentId, [new RepeaterRowCell { Span = 9 }]);

            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc));
            Assert.Contains("Segoe UI, 7 pt", result.Summary, StringComparison.Ordinal);
            var spacer = Control(doc, parentId).Descendants<TableRow>().Last();
            var mark = Assert.Single(spacer.Descendants<Paragraph>()).ParagraphProperties?.ParagraphMarkRunProperties;
            Assert.Equal(("Segoe UI", "14"), (Font(mark), Size(mark)));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Each_cell_takes_the_formatting_of_the_column_it_starts_in()
    {
        var path = Corpus.UnlinkedCopy(Corpus.SalesInvoice);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var repeater = LayoutEditor.InsertRepeaterTable(
                doc,
                "/Header/Line",
                TableArgs.DomainFieldsOf("/Header/Line", "ItemNo_Line", "Description_Line", "Quantity_Line"),
                new Location { Type = LocationKind.DocumentEnd },
                new RepeaterTableOptions());

            // The tool-built data row is unformatted; give its first column 8 pt and its last 10 pt bold.
            var dataRow = Control(doc, repeater.ControlId).Descendants<TableRow>().Single();
            var dataCells = dataRow.Elements<TableCell>().ToList();
            dataCells[0].Descendants<SdtRun>().Single().SdtProperties!
                .PrependChild(new RunProperties(new FontSize { Val = "16" }));
            dataCells[2].Descendants<SdtRun>().Single().SdtProperties!
                .PrependChild(new RunProperties(new Bold(), new FontSize { Val = "20" }));

            var result = LayoutEditor.InsertSubtotalRow(doc, repeater.ControlId,
            [
                new RepeaterRowCell { Span = 2, Fields = ["/Header/Line/Description_Line_Lbl"] },
                new RepeaterRowCell { Fields = ["/Header/Line/AmountExcludingVAT_Line"], Alignment = "right" },
            ]);

            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc));
            Assert.Contains("(8 pt; 10 pt, bold)", result.Summary, StringComparison.Ordinal);
            var cells = Control(doc, repeater.ControlId).Descendants<TableRow>().Last().Elements<TableCell>().ToList();
            var label = cells[0].Descendants<SdtRun>().Single().SdtProperties!.GetFirstChild<RunProperties>();
            var amount = cells[1].Descendants<SdtRun>().Single().SdtProperties!.GetFirstChild<RunProperties>();
            Assert.Equal("16", Size(label));
            Assert.False(IsBold(label));
            Assert.Equal("20", Size(amount));
            Assert.True(IsBold(amount));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_row_at_the_top_of_a_table_takes_the_row_below_it()
    {
        var path = CopyOfCorpus(Corpus.JobQuote);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var table = Control(doc, RepeaterId(doc, "/Job/Job_Task")).Ancestors<Table>().First();
            var tableIndex = TableGridNavigator.Tables(doc.MainDocumentPart!.Document!.Body!).ToList().IndexOf(table);

            // Job Quote's caption row is Segoe UI 8 pt, bold and underlined.
            var result = TableStructureEditor.InsertStaticRow(
                doc, LayoutPart.Body, null, tableIndex, atRow: 0,
                [new RepeaterRowCell { Span = TableGridNavigator.GridColumnCount(table) }]);

            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc));
            Assert.Contains("the cell below it (Segoe UI, 8 pt, bold)", result.Summary, StringComparison.Ordinal);
            var mark = TableGridNavigator.Rows(table)[0].InnerRow!.Descendants<Paragraph>().Single()
                .ParagraphProperties?.ParagraphMarkRunProperties;
            Assert.Equal(("Segoe UI", "16"), (Font(mark), Size(mark)));
            Assert.True(IsBold(mark));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_row_in_the_VAT_specification_invoice_takes_its_columns_paragraph_styles()
    {
        // StandardSalesInvoiceVatSpec sizes its line text through paragraph styles: grid column 0 is LeftAlign,
        // 2 RightAlign (which carries the 10 pt), 3 LeftAlign. A row built without them printed a point larger (P22).
        var path = CopyOfCorpus(Corpus.SalesInvoiceVatSpec);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var lines = RepeaterId(doc, "/Header/Line");
            var grid = TableGridNavigator.GridColumnCount(Control(doc, lines).Ancestors<Table>().First());

            var result = LayoutEditor.InsertSubtotalRow(doc, lines,
            [
                new RepeaterRowCell { Span = 2 },
                new RepeaterRowCell { Fields = ["/Header/Line/Quantity_Line"] },
                new RepeaterRowCell { Fields = ["/Header/Line/Description_Line"], Alignment = "right", Span = grid - 3 },
            ]);

            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc));
            Assert.Contains("paragraph style 'LeftAlign', 'RightAlign'", result.Summary, StringComparison.Ordinal);
            var paragraphs = Control(doc, lines).Descendants<TableRow>().Last()
                .Elements<TableCell>().Select(c => Assert.Single(c.Elements<Paragraph>())).ToList();
            Assert.Equal(["LeftAlign", "RightAlign", "LeftAlign"],
                paragraphs.Select(p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value).ToList());

            // The cell spec's own alignment is direct formatting, so it wins over the LeftAlign style.
            Assert.Equal(JustificationValues.Right, paragraphs[2].ParagraphProperties?.Justification?.Val?.Value);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_table_whose_rows_carry_no_direct_formatting_gives_none()
    {
        var path = Corpus.UnlinkedCopy(Corpus.SalesInvoice);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var repeater = LayoutEditor.InsertRepeaterTable(
                doc,
                "/Header/Line",
                TableArgs.DomainFieldsOf("/Header/Line", "ItemNo_Line", "Description_Line", "Quantity_Line"),
                new Location { Type = LocationKind.DocumentEnd },
                new RepeaterTableOptions());

            var result = LayoutEditor.InsertRepeaterRow(doc, repeater.ControlId, "/Header/Line/AssemblyLine",
                [new RepeaterRowCell { Span = 3, Fields = ["/Header/Line/AssemblyLine/Description_AssemblyLine"] }]);

            Assert.DoesNotContain("takes the formatting", result.Summary, StringComparison.Ordinal);
            var row = Control(doc, result.ControlId).Descendants<TableRow>().Single();
            Assert.Null(row.Descendants<SdtRun>().Single().SdtProperties!.GetFirstChild<RunProperties>());
            Assert.Null(row.Descendants<Paragraph>().Single().ParagraphProperties?.ParagraphMarkRunProperties);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertRepeaterRow_tool_takes_format()
    {
        var path = Corpus.UnlinkedCopy(Corpus.SalesInvoice);
        try
        {
            var table = TableTools.InsertRepeaterTable(
                path, "/Header/Line", TableArgs.FieldsOf("/Header/Line", "ItemNo_Line", "Description_Line", "Quantity_Line"), "documentEnd");
            Assert.True(table.Ok, table.Error?.Message);
            var tableDto = Assert.IsType<EditResultDto>(table.Data);

            var detail = TableTools.InsertRepeaterRow(
                path, tableDto.ControlId, "/Header/Line/AssemblyLine",
                [new() { Span = 3, Fields = ["/Header/Line/AssemblyLine/Description_AssemblyLine"] }],
                format: new TextFormatArgument { Italic = true, FontSizePoints = 8 });
            Assert.True(detail.Ok, detail.Error?.Message);
            var dto = Assert.IsType<EditResultDto>(detail.Data);

            using var doc = WordprocessingDocument.Open(path, false);
            var own = Control(doc, dto.ControlId).Descendants<SdtRun>().Single().SdtProperties!.GetFirstChild<RunProperties>();
            Assert.NotNull(own?.GetFirstChild<Italic>());
            Assert.Equal("16", Size(own));
        }
        finally
        {
            File.Delete(path);
        }
    }
}
