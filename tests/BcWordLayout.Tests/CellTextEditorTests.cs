using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

/// <summary>
/// Covers <see cref="CellTextEditor"/> directly (no MCP tool layer): each test opens a temp synthetic/
/// corpus layout editable, applies one plain-text cell operation, saves, then reopens read-only to assert
/// the on-disk result — mirroring the round-trip style of <c>LayoutEditorTests</c>.
/// </summary>
public class CellTextEditorTests
{
    private static Location Cell(int tableIndex, int row, int col) => new()
    {
        Type = LocationKind.TableCell,
        TableIndex = tableIndex,
        Row = row,
        Col = col,
    };

    [Fact]
    public void SetCellText_replaces_the_cells_text_and_stays_valid()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(rows: 1, cols: 2));
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var result = CellTextEditor.SetCellText(doc, Cell(0, 0, 0), "New Header");

                Assert.Equal("set_cell_text", result.Operation);
                Assert.Equal("R0C0", result.PreviousText);
                Assert.Equal("New Header", result.NewText);
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var firstCell = reopened.MainDocumentPart!.Document!.Descendants<TableCell>().First();
            Assert.Equal("New Header", string.Concat(firstCell.Descendants<Text>().Select(t => t.Text)));
            // The sibling cell is untouched.
            var secondCell = reopened.MainDocumentPart!.Document!.Descendants<TableCell>().ElementAt(1);
            Assert.Equal("R0C1", string.Concat(secondCell.Descendants<Text>().Select(t => t.Text)));

            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void SetCellText_preserves_the_cells_tcPr_and_paragraph_style()
    {
        // A cell carrying a w:tcPr (width) and a pStyle: both must survive a re-label.
        const string body =
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid>"
            + "<w:tr><w:tc><w:tcPr><w:tcW w:w=\"1234\" w:type=\"dxa\"/></w:tcPr>"
            + "<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr><w:r><w:t>Old</w:t></w:r></w:p>"
            + "</w:tc></w:tr></w:tbl>";
        var path = SyntheticLayout.Create(body);
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                CellTextEditor.SetCellText(doc, Cell(0, 0, 0), "New");
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var cell = reopened.MainDocumentPart!.Document!.Descendants<TableCell>().Single();
            Assert.Equal("1234", cell.TableCellProperties?.TableCellWidth?.Width?.Value);
            Assert.Equal("Heading1", cell.Descendants<ParagraphStyleId>().Single().Val?.Value);
            Assert.Equal("New", string.Concat(cell.Descendants<Text>().Select(t => t.Text)));

            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void SetCellText_optional_format_applies_bold_alignment_and_size()
    {
        // The knobs exist for cells in a freshly authored plain table (insert_table), which have no
        // styling to inherit - caption cells want bold, amount columns right alignment, titles a size.
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(rows: 1, cols: 2));
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                CellTextEditor.SetCellText(doc, Cell(0, 0, 0), "Order Confirmation",
                    new CellTextFormat { Bold = true, Alignment = "right", FontSizePoints = 16 });
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var cell = reopened.MainDocumentPart!.Document!.Descendants<TableCell>().First();
            var paragraph = cell.GetFirstChild<Paragraph>()!;
            Assert.Equal(JustificationValues.Right, paragraph.ParagraphProperties!.Justification!.Val!.Value);
            var rPr = paragraph.GetFirstChild<Run>()!.RunProperties!;
            Assert.NotNull(rPr.Bold);
            Assert.Equal("32", rPr.FontSize!.Val!.Value); // 16pt = 32 half-points
            Assert.Equal("32", rPr.FontSizeComplexScript!.Val!.Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void SetCellText_format_bold_false_strips_an_existing_bold()
    {
        var body =
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid>"
            + "<w:tr><w:tc><w:tcPr/><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Bold text</w:t></w:r></w:p></w:tc></w:tr></w:tbl>";
        var path = SyntheticLayout.Create(body);
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                CellTextEditor.SetCellText(doc, Cell(0, 0, 0), "Plain now", new CellTextFormat { Bold = false });
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var run = reopened.MainDocumentPart!.Document!.Descendants<TableCell>().First()
                .GetFirstChild<Paragraph>()!.GetFirstChild<Run>()!;
            Assert.Null(run.RunProperties?.Bold);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void SetCellText_without_format_still_preserves_existing_styling_untouched()
    {
        // The formatting knobs must be strictly additive: an omitted format leaves the old
        // preserve-everything behavior byte-for-byte (existing bold survives a plain re-label).
        var body =
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid>"
            + "<w:tr><w:tc><w:tcPr/><w:p><w:pPr><w:jc w:val=\"right\"/></w:pPr>"
            + "<w:r><w:rPr><w:b/></w:rPr><w:t>Old</w:t></w:r></w:p></w:tc></w:tr></w:tbl>";
        var path = SyntheticLayout.Create(body);
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                CellTextEditor.SetCellText(doc, Cell(0, 0, 0), "New");
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var paragraph = reopened.MainDocumentPart!.Document!.Descendants<TableCell>().First()
                .GetFirstChild<Paragraph>()!;
            Assert.Equal(JustificationValues.Right, paragraph.ParagraphProperties!.Justification!.Val!.Value);
            Assert.NotNull(paragraph.GetFirstChild<Run>()!.RunProperties!.Bold);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void SetCellText_collapses_a_multi_paragraph_cell_into_one()
    {
        // Mirrors the corpus "Amount" / "(ex. GST)" header: two paragraphs, several runs.
        const string body =
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid>"
            + "<w:tr><w:tc><w:tcPr/>"
            + "<w:p><w:r><w:t>Amount</w:t></w:r></w:p>"
            + "<w:p><w:r><w:t xml:space=\"preserve\">(ex. </w:t></w:r><w:r><w:t>GST</w:t></w:r><w:r><w:t>)</w:t></w:r></w:p>"
            + "</w:tc></w:tr></w:tbl>";
        var path = SyntheticLayout.Create(body);
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var result = CellTextEditor.SetCellText(doc, Cell(0, 0, 0), "Amount");
                Assert.Equal("Amount(ex. GST)", result.PreviousText);
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var cell = reopened.MainDocumentPart!.Document!.Descendants<TableCell>().Single();
            Assert.Single(cell.Elements<Paragraph>());
            Assert.Equal("Amount", string.Concat(cell.Descendants<Text>().Select(t => t.Text)));

            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void ClearCellText_empties_the_cell_but_leaves_a_valid_paragraph()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(rows: 1, cols: 2));
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var result = CellTextEditor.ClearCellText(doc, Cell(0, 0, 1));

                Assert.Equal("clear_cell_text", result.Operation);
                Assert.Equal("R0C1", result.PreviousText);
                Assert.Equal(string.Empty, result.NewText);
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var cells = reopened.MainDocumentPart!.Document!.Descendants<TableCell>().ToList();
            // Cleared cell: no text, but still a valid (paragraph-bearing) cell so the column stays.
            Assert.Equal(string.Empty, string.Concat(cells[1].Descendants<Text>().Select(t => t.Text)));
            Assert.NotEmpty(cells[1].Elements<Paragraph>());
            // Sibling untouched.
            Assert.Equal("R0C0", string.Concat(cells[0].Descendants<Text>().Select(t => t.Text)));

            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void SetCellText_on_a_cell_containing_a_content_control_throws_ArgumentException()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.TableWithBlockControlInCell(id: 88, text: "BOUND"));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var ex = Assert.Throws<ArgumentException>(() => CellTextEditor.SetCellText(doc, Cell(0, 0, 0), "x"));
            Assert.Contains("content control", ex.Message);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void ClearCellText_on_a_cell_level_control_cell_throws_ArgumentException()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.TableWithCellLevelControl(id: 99, text: "BOUND"));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var ex = Assert.Throws<ArgumentException>(() => CellTextEditor.ClearCellText(doc, Cell(0, 0, 0)));
            Assert.Contains("content control", ex.Message);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void SetCellText_with_a_non_tableCell_location_throws_ArgumentException()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(rows: 1, cols: 1));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            Assert.Throws<ArgumentException>(() =>
                CellTextEditor.SetCellText(doc, new Location { Type = LocationKind.DocumentEnd }, "x"));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void SetCellText_with_an_out_of_range_cell_throws_NotFoundException_with_TableCoordinate_target()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(rows: 1, cols: 2));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var ex = Assert.Throws<NotFoundException>(() => CellTextEditor.SetCellText(doc, Cell(0, 0, 5), "x"));
            Assert.Equal(NotFoundTarget.TableCoordinate, ex.TargetKind);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
