using BcWordLayout.Domain;
using BcWordLayout.McpHost;
using BcWordLayout.McpHost.Tools;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

/// <summary>
/// Covers <see cref="LayoutEditor.InsertPlainTable"/> and its <c>insert_table</c> tool surface — the
/// building block for authoring the NON-repeating sections of a layout from scratch (address columns,
/// header-info grids, totals blocks), born from the 2026-07-31 from-scratch authoring exercise where no
/// tool could produce them. The key contract: the result's TableIndex must be immediately addressable by
/// set_cell_text / insert_field (locationType 'tableCell').
/// </summary>
public class PlainTableTests
{
    private static string CopyOfCorpus(string corpusFile)
    {
        var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"bcwl-plaintable-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    private static List<ValidationErrorInfo> OpenXmlErrors(WordprocessingDocument doc) =>
        new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc).ToList();

    [Fact]
    public void Tool_insert_table_appends_an_empty_grid_the_follow_up_cell_tools_can_address()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = TableTools.InsertTable(path, rows: 2, columns: 3, locationType: "documentEnd",
                columnWidths: "3000,4000,3206", columnAlignments: "left,center,right");
            Assert.True(response.Ok, response.Error?.Message);
            var dto = Assert.IsType<TableEditResultDto>(response.Data);
            Assert.Equal("insert_table", dto.Operation);
            Assert.Equal(0, dto.ColumnCountBefore);
            Assert.Equal(3, dto.ColumnCountAfter);
            Assert.Equal(2, dto.RowsAffected);
            Assert.True(dto.QuickValidation.Passed, "fresh plain table must not introduce findings");

            // The returned index must address the new table for the follow-up cell tools.
            var setResponse = EditTools.SetCellText(path, dto.TableIndex, 0, 0, "Caption");
            Assert.True(setResponse.Ok, setResponse.Error?.Message);

            using var reopened = WordprocessingDocument.Open(path, false);
            var table = TableGridNavigator.Tables(reopened.MainDocumentPart!.Document!.Body!)[dto.TableIndex];
            var widths = table.GetFirstChild<TableGrid>()!.Elements<GridColumn>().Select(g => g.Width!.Value).ToList();
            Assert.Equal(new[] { "3000", "4000", "3206" }, widths);
            Assert.Null(table.GetFirstChild<TableProperties>()!.GetFirstChild<TableBorders>()); // borderless default

            var firstRowCells = table.Elements<TableRow>().First().Elements<TableCell>().ToList();
            Assert.Equal(3, firstRowCells.Count);
            Assert.Equal("Caption", firstRowCells[0].InnerText);
            // Per-column alignment seeds each cell's paragraph justification.
            Assert.Equal(
                JustificationValues.Center,
                firstRowCells[1].GetFirstChild<Paragraph>()!.ParagraphProperties!.Justification!.Val!.Value);
            Assert.Equal(
                JustificationValues.Right,
                firstRowCells[2].GetFirstChild<Paragraph>()!.ParagraphProperties!.Justification!.Val!.Value);

            Assert.Empty(TableGridConsistencyGuard.Find(reopened));
            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Tool_insert_table_borders_grid_draws_the_explicit_single_line_grid()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = TableTools.InsertTable(path, rows: 1, columns: 2, locationType: "documentEnd",
                borders: "grid");
            Assert.True(response.Ok, response.Error?.Message);
            var dto = Assert.IsType<TableEditResultDto>(response.Data);

            using var reopened = WordprocessingDocument.Open(path, false);
            var table = TableGridNavigator.Tables(reopened.MainDocumentPart!.Document!.Body!)[dto.TableIndex];
            var borders = table.GetFirstChild<TableProperties>()!.GetFirstChild<TableBorders>();
            Assert.NotNull(borders?.InsideHorizontalBorder);
            // Default widths: an even split of the full content width.
            Assert.Equal(
                new[] { "5103", "5103" },
                table.GetFirstChild<TableGrid>()!.Elements<GridColumn>().Select(g => g.Width!.Value));
            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Tool_insert_table_borders_headerRule_draws_set_cell_formats_bottom_rule_under_the_first_row_only()
    {
        // The repeater table's BC-native look, now on a plain table too (GitHub issue #131): one rule under row 0,
        // byte-identical to the one set_cell_format draws for a bottom edge, and nothing anywhere else.
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var ruled = Assert.IsType<TableEditResultDto>(TableTools.InsertTable(
                path, rows: 3, columns: 2, locationType: "documentEnd", borders: "headerRule").Data);
            var drawnLater = Assert.IsType<TableEditResultDto>(TableTools.InsertTable(
                path, rows: 3, columns: 2, locationType: "documentEnd").Data);
            Assert.True(TableTools.SetCellFormat(path, drawnLater.TableIndex, row: 0, edges: ["bottom"]).Ok);

            using var reopened = WordprocessingDocument.Open(path, false);
            var tables = TableGridNavigator.Tables(reopened.MainDocumentPart!.Document!.Body!);
            var table = tables[ruled.TableIndex];
            Assert.Null(table.GetFirstChild<TableProperties>()!.GetFirstChild<TableBorders>());

            var rows = table.Elements<TableRow>().ToList();
            var expected = tables[drawnLater.TableIndex].Elements<TableRow>().First()
                .Elements<TableCell>().Select(c => c.TableCellProperties!.TableCellBorders!.OuterXml).ToList();
            Assert.Equal(expected, rows[0].Elements<TableCell>().Select(c => c.TableCellProperties!.TableCellBorders!.OuterXml));
            Assert.All(rows.Skip(1).SelectMany(r => r.Elements<TableCell>()), c => Assert.Null(c.TableCellProperties!.TableCellBorders));

            // A rule, not a header: no w:tblHeader, which insert_column reads as "this row takes a header cell".
            Assert.Null(rows[0].TableRowProperties);
            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Tool_insert_table_tableStyle_references_a_table_style_the_layout_defines()
    {
        // insert_table had no tableStyle, so a body block could not take a branded shell's table style while
        // authoring (GitHub issue #131); it now takes the same option, with the same refusal for a style the
        // layout does not define.
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var dto = Assert.IsType<TableEditResultDto>(TableTools.InsertTable(
                path, rows: 1, columns: 2, locationType: "documentEnd", tableStyle: "TableGrid").Data);

            var refused = TableTools.InsertTable(path, rows: 1, columns: 2, locationType: "documentEnd", tableStyle: "NoSuchStyle");
            Assert.Equal("invalid_argument", refused.Error?.Code);
            Assert.Contains("'TableGrid'", refused.Error!.Message, StringComparison.Ordinal);

            using var reopened = WordprocessingDocument.Open(path, false);
            var tblPr = TableGridNavigator.Tables(reopened.MainDocumentPart!.Document!.Body!)[dto.TableIndex]
                .GetFirstChild<TableProperties>()!;
            Assert.Equal("TableGrid", tblPr.TableStyle!.Val!.Value);
            Assert.Equal(new[] { "tblStyle", "tblW", "tblLayout" }, tblPr.ChildElements.Select(e => e.LocalName));
            Assert.DoesNotContain(LayoutValidator.Quick(reopened).Findings, f => f.Check == "table-style-resolves");
            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("left,right", 3, "columnAlignments")]     // count mismatch (3 columns, 2 alignments)
    [InlineData("left,middle,right", 3, "columnAlignments")] // unknown alignment name
    public void Tool_insert_table_rejects_bad_alignment_input(string alignments, int columns, string expectedParam)
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = TableTools.InsertTable(path, rows: 1, columns: columns,
                locationType: "documentEnd", columnAlignments: alignments);
            Assert.False(response.Ok);
            Assert.Equal("invalid_argument", response.Error!.Code);
            Assert.Contains(expectedParam, response.Error.Message);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Tool_insert_table_in_the_default_header_lands_there_and_its_cells_take_text_in_that_part()
    {
        // Every stock header and footer is built from plain tables - the logo beside the document details, the
        // company details across the footer - so insert_table works there like insert_field (GitHub issue #132).
        // StandardPurchaseOrder's DEFAULT header is header2.xml, not the first header in the package, and the
        // layout renders a distinct first page: the summary has to say both.
        var path = CopyOfCorpus(Corpus.StandardPurchaseOrder);
        try
        {
            var response = TableTools.InsertTable(path, rows: 1, columns: 2, locationType: "documentEnd", layoutPart: "header");
            Assert.True(response.Ok, response.Error?.Message);
            var dto = Assert.IsType<TableEditResultDto>(response.Data);
            Assert.Equal("header2.xml", dto.Part);
            Assert.Contains("layoutPart 'header', partName 'header2.xml'", dto.Summary, StringComparison.Ordinal);
            Assert.Contains("DIFFERENT FIRST PAGE", dto.Summary, StringComparison.Ordinal);

            var filled = EditTools.SetCellText(path, dto.TableIndex, row: 0, col: 1, "Page", layoutPart: "header", partName: "header2.xml");
            Assert.True(filled.Ok, filled.Error?.Message);

            using var reopened = WordprocessingDocument.Open(path, false);
            var header = reopened.MainDocumentPart!.HeaderParts
                .Single(p => p.Uri.OriginalString.EndsWith("/header2.xml", StringComparison.Ordinal)).Header!;
            var table = TableGridNavigator.Tables(header)[dto.TableIndex];
            Assert.Same(header.Elements<Table>().Last(), table);
            Assert.IsType<Paragraph>(table.NextSibling());
            Assert.Equal("Page", table.Elements<TableRow>().Single().Elements<TableCell>().ElementAt(1).InnerText);
            Assert.Empty(TableGridConsistencyGuard.Find(reopened));
            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Tool_insert_table_in_a_blank_layouts_empty_footer_takes_its_lone_paragraph_as_the_trailing_one()
    {
        // create_layout scaffolds each header/footer with one empty paragraph (a part must hold a block). The
        // table goes BEFORE it, so that paragraph becomes the one InsertBlock guarantees after a table, and the
        // footer does not start with a stray empty line (the GitHub issue #91 shape, for block content).
        var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"bcwl-plaintable-blank-{Guid.NewGuid():N}.docx");
        try
        {
            LayoutBuilder.Create(Corpus.Path(Corpus.SalesInvoice), path);

            var dto = Assert.IsType<TableEditResultDto>(TableTools.InsertTable(
                path, rows: 1, columns: 3, locationType: "documentEnd", layoutPart: "footer").Data);
            Assert.Equal(0, dto.TableIndex);

            using var reopened = WordprocessingDocument.Open(path, false);
            var footer = reopened.MainDocumentPart!.FooterParts
                .Single(p => p.Uri.OriginalString.EndsWith("/" + dto.Part, StringComparison.Ordinal)).Footer!;
            Assert.Equal(new[] { "tbl", "p" }, footer.ChildElements.Select(e => e.LocalName));
            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Without_columnWidths_a_plain_table_placed_in_a_cell_splits_the_host_cells_width()
    {
        // A table nested in a cell has the cell to fit in, not the page: a page-wide fixed grid there overran the
        // host cell, and the outer table's own fixed grid with it. insert_repeater_table already split the host
        // cell's width (GitHub issue #84); insert_table now resolves its location first and does the same.
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var outer = Assert.IsType<TableEditResultDto>(TableTools.InsertTable(
                path, rows: 1, columns: 2, locationType: "documentEnd", columnWidths: "6000,4206").Data);

            var inner = Assert.IsType<TableEditResultDto>(TableTools.InsertTable(
                path, rows: 1, columns: 2, locationType: "tableCell", tableIndex: outer.TableIndex, row: 0, col: 0).Data);

            using var reopened = WordprocessingDocument.Open(path, false);
            var innerTable = TableGridNavigator.Tables(reopened.MainDocumentPart!.Document!.Body!)[inner.TableIndex];
            Assert.Equal(
                new[] { "3000", "3000" },
                innerTable.GetFirstChild<TableGrid>()!.Elements<GridColumn>().Select(g => g.Width!.Value));
            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Inserted_control_takes_optional_bold_and_size_on_its_runs_and_sdtPr()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var tableResponse = TableTools.InsertTable(path, rows: 1, columns: 2, locationType: "documentEnd");
            var dto = Assert.IsType<TableEditResultDto>(tableResponse.Data);

            var response = EditTools.InsertLabel(
                path, "/Header/ExternalDocumentNo_Lbl", "tableCell",
                tableIndex: dto.TableIndex, row: 0, col: 0, format: new TextFormatArgument { Bold = true, FontSizePoints = 8 });
            Assert.True(response.Ok, response.Error?.Message);

            using var reopened = WordprocessingDocument.Open(path, false);
            var table = TableGridNavigator.Tables(reopened.MainDocumentPart!.Document!.Body!)[dto.TableIndex];
            var sdt = table.Descendants<SdtRun>().Single();
            var runProps = sdt.Descendants<Run>().Single().RunProperties!;
            Assert.NotNull(runProps.Bold);
            Assert.Equal("16", runProps.FontSize!.Val!.Value); // 8pt = 16 half-points
            // sdtPr rPr styles the placeholder/future content; corpus order puts it first.
            var sdtPrRunProps = sdt.SdtProperties!.GetFirstChild<RunProperties>();
            Assert.NotNull(sdtPrRunProps?.Bold);
            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Stacked_block_tables_get_a_separator_paragraph_so_Word_never_merges_them()
    {
        // Word renders two adjacent w:tbl siblings as ONE merged table; authoring in Word always keeps a
        // paragraph between them. Stacking blocks (address grid straight after a title table, a repeater
        // straight after an info grid) must therefore never leave two tables touching.
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            Assert.True(TableTools.InsertTable(path, rows: 1, columns: 2, locationType: "documentEnd").Ok);
            Assert.True(TableTools.InsertTable(path, rows: 1, columns: 3, locationType: "documentEnd").Ok);
            Assert.True(TableTools.InsertRepeaterTable(
                path, "/Header/Line", "ItemNo_Line,Description_Line", "documentEnd").Ok);

            using var reopened = WordprocessingDocument.Open(path, false);
            var body = reopened.MainDocumentPart!.Document!.Body!;
            var adjacentPairs = body.ChildElements.Zip(body.ChildElements.Skip(1))
                .Count(pair => pair.First is Table && pair.Second is Table);
            Assert.Equal(0, adjacentPairs);
            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Fresh_plain_table_cell_takes_a_bound_field_via_tableCell_location()
    {
        // The authoring loop this tool exists for: insert_table -> insert_field into one of its cells.
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var tableResponse = TableTools.InsertTable(path, rows: 1, columns: 2, locationType: "documentEnd");
            var dto = Assert.IsType<TableEditResultDto>(tableResponse.Data);

            var fieldResponse = EditTools.InsertField(
                path, "/Header/ExternalDocumentNo", "tableCell", tableIndex: dto.TableIndex, row: 0, col: 1);
            Assert.True(fieldResponse.Ok, fieldResponse.Error?.Message);

            using var reopened = WordprocessingDocument.Open(path, false);
            var table = TableGridNavigator.Tables(reopened.MainDocumentPart!.Document!.Body!)[dto.TableIndex];
            var cell = table.Elements<TableRow>().First().Elements<TableCell>().ElementAt(1);
            Assert.Single(cell.Descendants<SdtRun>());
            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Fresh_plain_table_pins_its_grid_with_tblW_and_a_fixed_tblLayout()
    {
        // Without w:tblLayout the OOXML default is autofit, so Word recomputes every column from cell
        // content and the columnWidths passed here hold only until someone opens the file - one plain Word
        // save rewrote an untouched 2800,3000 grid to 3177,3066 (GitHub issue #52). Every corpus table
        // declares both elements.
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = TableTools.InsertTable(
                path, rows: 2, columns: 3, locationType: "documentEnd", columnWidths: "3000,4000,3206");
            Assert.True(response.Ok, response.Error?.Message);
            var dto = Assert.IsType<TableEditResultDto>(response.Data);

            using var reopened = WordprocessingDocument.Open(path, false);
            var table = TableGridNavigator.Tables(reopened.MainDocumentPart!.Document!.Body!)[dto.TableIndex];
            var tblPr = table.GetFirstChild<TableProperties>()!;

            Assert.Equal(TableLayoutValues.Fixed, tblPr.TableLayout!.Type!.Value);
            Assert.Equal(TableWidthUnitValues.Dxa, tblPr.TableWidth!.Type!.Value);

            // The declared total is the grid's own sum, so the two can never disagree.
            var gridSum = table.GetFirstChild<TableGrid>()!.Elements<GridColumn>()
                .Sum(c => int.Parse(c.Width!.Value!, System.Globalization.CultureInfo.InvariantCulture));
            Assert.Equal(3000 + 4000 + 3206, gridSum);
            Assert.Equal(gridSum.ToString(System.Globalization.CultureInfo.InvariantCulture), tblPr.TableWidth.Width!.Value);

            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Bordered_plain_table_orders_tblPr_children_as_the_schema_requires()
    {
        // CT_TblPrBase is a SEQUENCE: w:tblW precedes w:tblBorders, which precedes w:tblLayout. Appending
        // the new elements instead of assigning the SDK's typed properties would emit them out of order -
        // valid-looking C#, invalid OOXML - so this asserts the order explicitly rather than relying on
        // OpenXmlValidator alone to notice.
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = TableTools.InsertTable(
                path, rows: 1, columns: 2, locationType: "documentEnd", borders: "grid");
            var dto = Assert.IsType<TableEditResultDto>(response.Data);

            using var reopened = WordprocessingDocument.Open(path, false);
            var table = TableGridNavigator.Tables(reopened.MainDocumentPart!.Document!.Body!)[dto.TableIndex];
            var names = table.GetFirstChild<TableProperties>()!.ChildElements
                .Select(e => e.LocalName).ToList();

            Assert.Equal(new[] { "tblW", "tblBorders", "tblLayout" }, names);
            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }
}
