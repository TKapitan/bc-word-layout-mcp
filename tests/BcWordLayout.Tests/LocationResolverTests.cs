using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

public class LocationResolverTests
{
    private static int? ReadId(SdtElement sdt) =>
        sdt.GetFirstChild<SdtProperties>()?.GetFirstChild<SdtId>()?.Val?.Value;

    private static string CopyOfCorpus(string corpusFile)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-locationresolver-part-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    // ---- DocumentEnd: append at end of body, before the trailing w:sectPr ----

    [Fact]
    public void DocumentEnd_InsertBlock_lands_immediately_before_the_trailing_sectPr()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.PlainParagraph("existing"));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var body = doc.MainDocumentPart!.Document!.Body!;
            var sectPr = body.Elements<SectionProperties>().Single();

            var anchor = LocationResolver.Resolve(new Location { Type = LocationKind.DocumentEnd }, doc);
            var marker = new Paragraph(new Run(new Text("MARKER-BLOCK")));
            anchor.InsertBlock(marker);

            Assert.Same(marker, sectPr.PreviousSibling());
            Assert.Same(body, marker.Parent);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void DocumentEnd_InsertInline_wraps_a_new_paragraph_immediately_before_the_trailing_sectPr()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.PlainParagraph("existing"));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var body = doc.MainDocumentPart!.Document!.Body!;
            var sectPr = body.Elements<SectionProperties>().Single();

            var anchor = LocationResolver.Resolve(new Location { Type = LocationKind.DocumentEnd }, doc);
            var marker = new Run(new Text("MARKER-INLINE"));
            anchor.InsertInline(marker);

            var newParagraph = Assert.IsType<Paragraph>(sectPr.PreviousSibling());
            Assert.Same(marker, newParagraph.FirstChild);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void DocumentEnd_with_no_sectPr_appends_at_the_very_end_of_the_body()
    {
        // A body with no trailing sectPr at all (unusual, but Body doesn't strictly require one) —
        // InsertBlock must fall back to a plain append rather than throwing.
        var path = SyntheticLayout.Create(SyntheticLayout.PlainParagraph("existing"));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var body = doc.MainDocumentPart!.Document!.Body!;
            body.Elements<SectionProperties>().Single().Remove();

            var anchor = LocationResolver.Resolve(new Location { Type = LocationKind.DocumentEnd }, doc);
            var marker = new Paragraph(new Run(new Text("MARKER-BLOCK")));
            anchor.InsertBlock(marker);

            Assert.Same(marker, body.LastChild);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- AfterControl: find the sdt by w:id; anchor immediately after it, same parent ----

    [Fact]
    public void AfterControl_on_an_inline_control_inserts_inline_content_right_after_it_in_the_same_paragraph()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.InlineControlWithId(55));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var body = doc.MainDocumentPart!.Document!.Body!;
            var target = body.Descendants<SdtElement>().Single(s => ReadId(s) == 55);
            var paragraph = Assert.IsType<Paragraph>(target.Parent);

            var anchor = LocationResolver.Resolve(new Location { Type = LocationKind.AfterControl, ControlId = 55 }, doc);
            var marker = new Run(new Text("MARKER"));
            anchor.InsertInline(marker);

            Assert.Same(marker, target.NextSibling());
            Assert.Same(paragraph, marker.Parent);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void AfterControl_on_a_block_control_inserts_block_content_right_after_it_same_parent()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.BlockControlWithId(77));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var body = doc.MainDocumentPart!.Document!.Body!;
            var target = body.Descendants<SdtElement>().Single(s => ReadId(s) == 77);

            var anchor = LocationResolver.Resolve(new Location { Type = LocationKind.AfterControl, ControlId = 77 }, doc);
            var marker = new Paragraph(new Run(new Text("MARKER-BLOCK")));
            anchor.InsertBlock(marker);

            Assert.Same(marker, target.NextSibling());
            Assert.Same(body, marker.Parent);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void AfterControl_block_insert_after_an_inline_control_is_promoted_to_after_the_enclosing_paragraph()
    {
        // Block content (e.g. a repeater table) cannot be inserted as a sibling INSIDE a paragraph, so
        // targeting an inline control for a block insert must promote to right after that control's
        // enclosing paragraph instead of throwing or corrupting the document.
        var path = SyntheticLayout.Create(SyntheticLayout.InlineControlWithId(88));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var body = doc.MainDocumentPart!.Document!.Body!;
            var target = body.Descendants<SdtElement>().Single(s => ReadId(s) == 88);
            var paragraph = Assert.IsType<Paragraph>(target.Parent);

            var anchor = LocationResolver.Resolve(new Location { Type = LocationKind.AfterControl, ControlId = 88 }, doc);
            var marker = new Paragraph(new Run(new Text("MARKER-BLOCK")));
            anchor.InsertBlock(marker);

            Assert.Same(marker, paragraph.NextSibling());
            Assert.Same(body, marker.Parent);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void AfterControl_finds_the_right_control_among_several_by_id()
    {
        var body = SyntheticLayout.InlineControlWithId(1, "first")
            + SyntheticLayout.InlineControlWithId(2, "second")
            + SyntheticLayout.InlineControlWithId(3, "third");
        var path = SyntheticLayout.Create(body);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var docBody = doc.MainDocumentPart!.Document!.Body!;
            var target2 = docBody.Descendants<SdtElement>().Single(s => ReadId(s) == 2);

            var anchor = LocationResolver.Resolve(new Location { Type = LocationKind.AfterControl, ControlId = 2 }, doc);
            var marker = new Run(new Text("MARKER"));
            anchor.InsertInline(marker);

            Assert.Same(marker, target2.NextSibling());
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void AfterControl_with_unknown_id_throws_NotFoundException_with_Control_target()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.InlineControlWithId(1));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);

            var ex = Assert.Throws<NotFoundException>(() =>
                LocationResolver.Resolve(new Location { Type = LocationKind.AfterControl, ControlId = 9999 }, doc));
            Assert.Contains("9999", ex.Message);
            Assert.Equal(NotFoundTarget.Control, ex.TargetKind);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- AfterControl on a cell-level control (SdtCell, parent w:tr): must anchor INSIDE that
    // control's own cell, never as a sibling of the sdt in the row — a <w:tr> cannot host a paragraph or
    // block sdt directly. Real corpus controls in header tables (e.g. YourReference_Lbl) are this shape. ----

    [Fact]
    public void AfterControl_on_a_cell_level_control_anchors_inside_that_controls_own_cell_not_as_a_row_sibling()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.TableWithCellLevelControl(id: -1130623254, text: "YourReference_Lbl"));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var body = doc.MainDocumentPart!.Document!.Body!;
            var row = body.Descendants<TableRow>().Single();

            // The control's own cell lives INSIDE its SdtCell wrapper's SdtContentCell, not as a plain
            // w:tc sibling of the row (that's the whole point being tested); the "other" cell is the
            // row's one ordinary w:tc.
            var targetSdtCell = row.Elements<SdtCell>().Single();
            var targetCell = targetSdtCell.SdtContentCell!.GetFirstChild<TableCell>()!;
            var otherCell = row.Elements<TableCell>().Single();

            var anchor = LocationResolver.Resolve(
                new Location { Type = LocationKind.AfterControl, ControlId = -1130623254 }, doc);
            var marker = new Run(new Text("MARKER-AFTER-CELL"));
            anchor.InsertInline(marker);

            // Never a direct child of w:tr (that would be invalid OOXML - the bug this test guards against).
            Assert.IsNotType<TableRow>(marker.Parent);
            Assert.Same(targetCell, marker.Ancestors<TableCell>().First());
            Assert.Null(otherCell.Descendants<Run>().FirstOrDefault(r => ReferenceEquals(r, marker)));

            // Still exactly two cells in the row - InsertInline must not have added a new w:tc/sdt sibling.
            Assert.Equal(2, row.Elements<TableCell>().Count() + row.Elements<SdtCell>().Count());
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void AfterControl_on_a_row_level_control_throws_NotFoundException_with_AfterControlPosition_target()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.TableWithRowLevelControl(id: 287936249));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);

            var ex = Assert.Throws<NotFoundException>(() =>
                LocationResolver.Resolve(new Location { Type = LocationKind.AfterControl, ControlId = 287936249 }, doc));
            Assert.Contains("row-level", ex.Message);
            Assert.Contains("TableCell", ex.Message);
            Assert.Equal(NotFoundTarget.AfterControlPosition, ex.TargetKind);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void AfterControl_on_the_real_corpus_cell_level_label_inserts_a_field_that_passes_OpenXmlValidator()
    {
        // Regression test for the exact bug reported against a real corpus control: YourReference_Lbl
        // (id -1130623254) in SalesInvoiceForSubscriptionBilling.docx is a cell-level sdt (its sdtContent
        // wraps a whole <w:tc>, so its parent is a <w:tr>). AfterControl used to insert a wrapping <w:p>
        // straight into that <w:tr>, which OpenXmlValidator flags as an invalid child element.
        const int cellLevelLabelId = -1130623254;
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-aftercontrol-cell-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(Corpus.SalesInvoice), path, overwrite: true);

        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var body = doc.MainDocumentPart!.Document!.Body!;
                var labelSdt = body.Descendants<SdtElement>().Single(s => ReadId(s) == cellLevelLabelId);
                Assert.IsType<SdtCell>(labelSdt);
                Assert.IsType<TableRow>(labelSdt.Parent);

                var schema = SchemaProvider.FromLayout(doc);
                var field = SdtFactory.BuildField(schema, "/Header/SalesPersonName", placeholderText: "AFTER-CELL");

                var anchor = LocationResolver.Resolve(
                    new Location { Type = LocationKind.AfterControl, ControlId = cellLevelLabelId }, doc);
                anchor.InsertInline(field);

                doc.MainDocumentPart!.Document!.Save();
            }

            using (var reopened = WordprocessingDocument.Open(path, false))
            {
                var openXmlErrors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened).ToList();
                Assert.Empty(openXmlErrors);

                var body = reopened.MainDocumentPart!.Document!.Body!;
                var labelSdt = Assert.IsType<SdtCell>(body.Descendants<SdtElement>().Single(s => ReadId(s) == cellLevelLabelId));
                var cell = labelSdt.SdtContentCell!.GetFirstChild<TableCell>()!;

                // The newly-inserted field must be inside THAT SAME cell (not a new sibling row/cell).
                Assert.Contains(cell.Descendants<Text>(), t => t.Text == "AFTER-CELL");
            }
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- TableCell: (tableIndex, row, col), 0-based; anchor inside that cell ----

    [Fact]
    public void TableCell_InsertInline_lands_in_the_targeted_cells_own_paragraph()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(rows: 2, cols: 3));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var body = doc.MainDocumentPart!.Document!.Body!;
            var table = body.Descendants<Table>().Single();
            var targetCell = table.Elements<TableRow>().ElementAt(1).Elements<TableCell>().ElementAt(2);
            var otherCell = table.Elements<TableRow>().ElementAt(0).Elements<TableCell>().ElementAt(0);

            var anchor = LocationResolver.Resolve(
                new Location { Type = LocationKind.TableCell, TableIndex = 0, Row = 1, Col = 2 }, doc);
            var marker = new Run(new Text("MARKER-CELL"));
            anchor.InsertInline(marker);

            Assert.Same(targetCell, marker.Ancestors<TableCell>().First());

            // Not Assert.DoesNotContain(marker, ...): xUnit's default equality comparer for a T that
            // implements IEnumerable (every OpenXmlElement, including Run, enumerates its own children)
            // compares sequences structurally rather than by reference, so a different Run with the same
            // one-child shape can misleadingly compare "equal" to marker regardless of its actual text.
            // Assert.Null on an explicit ReferenceEquals lookup says exactly what's meant: marker itself
            // must not be among otherCell's runs.
            Assert.Null(otherCell.Descendants<Run>().FirstOrDefault(r => ReferenceEquals(r, marker)));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void TableCell_InsertBlock_lands_immediately_before_the_cells_trailing_paragraph()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(rows: 1, cols: 1));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var body = doc.MainDocumentPart!.Document!.Body!;
            var cell = body.Descendants<TableCell>().Single();
            var trailingParagraph = cell.Elements<Paragraph>().Last();

            var anchor = LocationResolver.Resolve(
                new Location { Type = LocationKind.TableCell, TableIndex = 0, Row = 0, Col = 0 }, doc);
            var marker = new Paragraph(new Run(new Text("MARKER-BLOCK-IN-CELL")));
            anchor.InsertBlock(marker);

            Assert.Same(marker, trailingParagraph.PreviousSibling());
            Assert.Same(cell, marker.Parent);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void TableCell_out_of_range_table_index_throws_NotFoundException_with_TableCoordinate_target()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(1, 1));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);

            var ex = Assert.Throws<NotFoundException>(() =>
                LocationResolver.Resolve(new Location { Type = LocationKind.TableCell, TableIndex = 5, Row = 0, Col = 0 }, doc));
            Assert.Contains("Table index", ex.Message);
            Assert.Equal(NotFoundTarget.TableCoordinate, ex.TargetKind);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void TableCell_out_of_range_row_throws_NotFoundException_with_TableCoordinate_target()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(2, 2));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);

            var ex = Assert.Throws<NotFoundException>(() =>
                LocationResolver.Resolve(new Location { Type = LocationKind.TableCell, TableIndex = 0, Row = 9, Col = 0 }, doc));
            Assert.Contains("Row index", ex.Message);
            Assert.Equal(NotFoundTarget.TableCoordinate, ex.TargetKind);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void TableCell_out_of_range_col_throws_NotFoundException_with_TableCoordinate_target()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(2, 2));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);

            var ex = Assert.Throws<NotFoundException>(() =>
                LocationResolver.Resolve(new Location { Type = LocationKind.TableCell, TableIndex = 0, Row = 0, Col = 9 }, doc));
            Assert.Contains("Column index", ex.Message);
            Assert.Equal(NotFoundTarget.TableCoordinate, ex.TargetKind);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- TableCell must count cell-level (SdtCell) and row-level (SdtRow) content-control wrappers as
    // cells/rows, EXACTLY as TableStructureReader (get_layout_info) does - otherwise a col/row index a
    // caller reads back from get_layout_info addresses a different physical spot (or wrongly reports out of
    // range). Regression for the reported bug: a BC header row of three address fields is
    // [SdtCell, w:tc, SdtCell]; counting only bare w:tc saw 1 cell, so col 1 failed as out of range. ----

    [Fact]
    public void TableCell_counts_a_cell_level_control_as_a_cell_so_its_plain_sibling_is_reachable_by_index()
    {
        // Row shape: [SdtCell(id 42) -> col 0, plain w:tc "plain-cell" -> col 1]. Before the fix,
        // row.Elements<TableCell>() saw only the 1 bare w:tc, so col 1 was reported out of range.
        var path = SyntheticLayout.Create(SyntheticLayout.TableWithCellLevelControl(id: 42, text: "cell-ctl"));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var body = doc.MainDocumentPart!.Document!.Body!;
            var row = body.Descendants<TableRow>().Single();
            var plainCell = row.Elements<TableCell>().Single();

            var anchor = LocationResolver.Resolve(
                new Location { Type = LocationKind.TableCell, TableIndex = 0, Row = 0, Col = 1 }, doc);
            var marker = new Run(new Text("MARKER-PLAIN-SIBLING"));
            anchor.InsertInline(marker);

            Assert.Same(plainCell, marker.Ancestors<TableCell>().First());
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void TableCell_targeting_a_cell_level_control_column_anchors_inside_that_controls_own_cell()
    {
        // col 0 is the SdtCell itself - it must resolve to the w:tc INSIDE its sdtContent (same cell the
        // AfterControl-on-cell-level path reaches), not fail to resolve.
        var path = SyntheticLayout.Create(SyntheticLayout.TableWithCellLevelControl(id: 42, text: "cell-ctl"));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var body = doc.MainDocumentPart!.Document!.Body!;
            var sdtCell = body.Descendants<SdtCell>().Single();
            var innerCell = sdtCell.SdtContentCell!.GetFirstChild<TableCell>()!;

            var anchor = LocationResolver.Resolve(
                new Location { Type = LocationKind.TableCell, TableIndex = 0, Row = 0, Col = 0 }, doc);
            var marker = new Run(new Text("MARKER-INSIDE-CELL-CTL"));
            anchor.InsertInline(marker);

            Assert.Same(innerCell, marker.Ancestors<TableCell>().First());
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void TableCell_counts_a_row_level_control_as_a_row_and_anchors_inside_its_inner_row_cell()
    {
        // The single "row" is a row-level SdtRow whose sdtContent wraps the data w:tr. Before the fix,
        // table.Elements<TableRow>() saw 0 rows, so row 0 was reported out of range.
        var path = SyntheticLayout.Create(SyntheticLayout.TableWithRowLevelControl(id: 99, text: "row-ctl"));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var body = doc.MainDocumentPart!.Document!.Body!;
            var innerCell = body.Descendants<SdtRow>().Single()
                .SdtContentRow!.GetFirstChild<TableRow>()!.Elements<TableCell>().Single();

            var anchor = LocationResolver.Resolve(
                new Location { Type = LocationKind.TableCell, TableIndex = 0, Row = 0, Col = 0 }, doc);
            var marker = new Run(new Text("MARKER-ROW-CTL-CELL"));
            anchor.InsertInline(marker);

            Assert.Same(innerCell, marker.Ancestors<TableCell>().First());
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void TableCell_on_the_real_corpus_address_row_resolves_the_empty_middle_cell_and_inserts_clean()
    {
        // The exact reported scenario: SalesInvoiceForSubscriptionBilling.docx table 0, row 7 is
        // [CustomerAddress8 (SdtCell) | empty w:tc] - two cells per get_layout_info. Inserting the DueDate
        // field into col 1 (the empty second cell) used to fail with "Column index 1 is out of range ...
        // it has 1 cell(s)"; it must now resolve and round-trip clean through OpenXmlValidator.
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-tablecell-corpus-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(Corpus.SalesInvoice), path, overwrite: true);
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var schema = SchemaProvider.FromLayout(doc);
                var field = SdtFactory.BuildField(schema, "/Header/DueDate", placeholderText: "DUEDATE-CELL");

                var anchor = LocationResolver.Resolve(
                    new Location { Type = LocationKind.TableCell, TableIndex = 0, Row = 7, Col = 1 }, doc);
                anchor.InsertInline(field);

                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var openXmlErrors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened).ToList();
            Assert.Empty(openXmlErrors);
            Assert.Contains(reopened.MainDocumentPart!.Document!.Body!.Descendants<Text>(), t => t.Text == "DUEDATE-CELL");
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- AtText: first w:t containing SearchText; anchor is its containing paragraph ----

    [Fact]
    public void AtText_InsertInline_appends_into_the_matching_paragraph()
    {
        var body = SyntheticLayout.PlainParagraph("before")
            + SyntheticLayout.PlainParagraph("find-ME-here-please")
            + SyntheticLayout.PlainParagraph("after");
        var path = SyntheticLayout.Create(body);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var docBody = doc.MainDocumentPart!.Document!.Body!;
            var targetParagraph = docBody.Descendants<Text>()
                .Single(t => t.Text.Contains("find-ME-here"))
                .Ancestors<Paragraph>()
                .First();

            var anchor = LocationResolver.Resolve(new Location { Type = LocationKind.AtText, SearchText = "find-ME-here" }, doc);
            var marker = new Run(new Text("MARKER-AT-TEXT"));
            anchor.InsertInline(marker);

            Assert.Same(targetParagraph, marker.Parent);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void AtText_InsertBlock_inserts_a_new_paragraph_immediately_after_the_matching_paragraph()
    {
        var body = SyntheticLayout.PlainParagraph("before")
            + SyntheticLayout.PlainParagraph("find-ME-here-please")
            + SyntheticLayout.PlainParagraph("after");
        var path = SyntheticLayout.Create(body);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var docBody = doc.MainDocumentPart!.Document!.Body!;
            var targetParagraph = docBody.Descendants<Text>()
                .Single(t => t.Text.Contains("find-ME-here"))
                .Ancestors<Paragraph>()
                .First();

            var anchor = LocationResolver.Resolve(new Location { Type = LocationKind.AtText, SearchText = "find-ME-here" }, doc);
            var marker = new Paragraph(new Run(new Text("MARKER-BLOCK-AT-TEXT")));
            anchor.InsertBlock(marker);

            Assert.Same(marker, targetParagraph.NextSibling());
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void AtText_text_not_found_throws_NotFoundException_with_SearchText_target()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.PlainParagraph("hello world"));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);

            var ex = Assert.Throws<NotFoundException>(() =>
                LocationResolver.Resolve(new Location { Type = LocationKind.AtText, SearchText = "no-such-text" }, doc));
            Assert.Contains("no-such-text", ex.Message);
            Assert.Equal(NotFoundTarget.SearchText, ex.TargetKind);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- Location.Validate(): required-field checks per kind ----

    [Fact]
    public void Validate_DocumentEnd_never_throws()
    {
        new Location { Type = LocationKind.DocumentEnd }.Validate();
    }

    [Fact]
    public void Validate_AfterControl_without_ControlId_throws_ArgumentException()
    {
        var location = new Location { Type = LocationKind.AfterControl };
        Assert.Throws<ArgumentException>(location.Validate);
    }

    [Fact]
    public void Validate_TableCell_missing_any_index_throws_ArgumentException()
    {
        Assert.Throws<ArgumentException>(new Location { Type = LocationKind.TableCell, Row = 0, Col = 0 }.Validate);
        Assert.Throws<ArgumentException>(new Location { Type = LocationKind.TableCell, TableIndex = 0, Col = 0 }.Validate);
        Assert.Throws<ArgumentException>(new Location { Type = LocationKind.TableCell, TableIndex = 0, Row = 0 }.Validate);
    }

    [Theory]
    [InlineData(-1, 0, 0)]
    [InlineData(0, -1, 0)]
    [InlineData(0, 0, -1)]
    public void Validate_TableCell_negative_index_throws_ArgumentException(int tableIndex, int row, int col)
    {
        var location = new Location { Type = LocationKind.TableCell, TableIndex = tableIndex, Row = row, Col = col };
        Assert.Throws<ArgumentException>(location.Validate);
    }

    [Fact]
    public void Validate_AtText_missing_SearchText_throws_ArgumentException()
    {
        Assert.Throws<ArgumentException>(new Location { Type = LocationKind.AtText }.Validate);
        Assert.Throws<ArgumentException>(new Location { Type = LocationKind.AtText, SearchText = "" }.Validate);
    }

    [Fact]
    public void Resolve_calls_Validate_so_a_structurally_invalid_location_throws_ArgumentException()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.PlainParagraph("x"));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);

            Assert.Throws<ArgumentException>(() =>
                LocationResolver.Resolve(new Location { Type = LocationKind.AfterControl }, doc));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Resolve_null_location_throws_ArgumentNullException()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.PlainParagraph("x"));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);

            Assert.Throws<ArgumentNullException>(() => LocationResolver.Resolve(null!, doc));
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- Location.Part / PartName (header/footer targeting) ----

    [Fact]
    public void Location_Part_defaults_to_Body()
    {
        // Regression guard: every pre-existing caller (this whole test file, LayoutEditor, every tool)
        // constructs a Location without ever setting Part, so the default must stay Body or every one of
        // them would silently start targeting something else.
        Assert.Equal(LayoutPart.Body, new Location { Type = LocationKind.DocumentEnd }.Part);
    }

    [Fact]
    public void DocumentEnd_with_explicit_Part_Body_behaves_identically_to_the_omitted_default()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.PlainParagraph("existing"));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var body = doc.MainDocumentPart!.Document!.Body!;
            var sectPr = body.Elements<SectionProperties>().Single();

            var anchor = LocationResolver.Resolve(new Location { Type = LocationKind.DocumentEnd, Part = LayoutPart.Body }, doc);
            Assert.Equal("document.xml", anchor.PartName);

            var marker = new Paragraph(new Run(new Text("MARKER-EXPLICIT-BODY")));
            anchor.InsertBlock(marker);

            Assert.Same(marker, sectPr.PreviousSibling());
            Assert.Same(body, marker.Parent);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void DocumentEnd_on_Header_targets_the_section_default_header_and_appends_at_its_own_end()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var main = doc.MainDocumentPart!;

            // This test used to derive its expectation from main.HeaderParts.First(), which made it agree
            // with the resolver by construction whatever either of them did. It now names the part
            // independently: header1.xml is the DEFAULT header of this layout's only section, while
            // HeaderParts.First() is header2.xml - so the old form would still pass against the old
            // positional rule that put content in the wrong header.
            var anchor = LocationResolver.Resolve(new Location { Type = LocationKind.DocumentEnd, Part = LayoutPart.Header }, doc);
            Assert.Equal("header1.xml", anchor.PartName);

            var target = main.HeaderParts.Single(h => Path.GetFileName(h.Uri.OriginalString) == "header1.xml");
            var marker = new Paragraph(new Run(new Text("MARKER-HEADER-END")));
            anchor.InsertBlock(marker);

            // Landed at the very end of that header's own content (a header never has a w:sectPr to insert
            // before, so this is a plain append) - not in the body, not in any other header/footer.
            Assert.Same(marker, target.Header!.LastChild);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void DocumentEnd_on_Footer_targets_the_section_default_footer_and_appends_at_its_own_end()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var main = doc.MainDocumentPart!;

            // Named independently rather than derived from FooterParts.First() - see the header test above.
            var anchor = LocationResolver.Resolve(new Location { Type = LocationKind.DocumentEnd, Part = LayoutPart.Footer }, doc);
            Assert.Equal("footer1.xml", anchor.PartName);

            var target = main.FooterParts.Single(f => Path.GetFileName(f.Uri.OriginalString) == "footer1.xml");
            var marker = new Paragraph(new Run(new Text("MARKER-FOOTER-END")));
            anchor.InsertBlock(marker);

            Assert.Same(marker, target.Footer!.LastChild);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Header_with_explicit_PartName_targets_that_specific_part_not_the_first()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var headers = doc.MainDocumentPart!.HeaderParts.ToList();
            Assert.True(headers.Count >= 2, "test assumes SalesInvoiceForSubscriptionBilling.docx has at least 2 header parts");

            var second = headers[1];
            var secondName = Path.GetFileName(second.Uri.OriginalString);

            var anchor = LocationResolver.Resolve(
                new Location { Type = LocationKind.DocumentEnd, Part = LayoutPart.Header, PartName = secondName }, doc);
            Assert.Equal(secondName, anchor.PartName);

            var marker = new Paragraph(new Run(new Text("MARKER-SECOND-HEADER")));
            anchor.InsertBlock(marker);

            Assert.Same(marker, second.Header!.LastChild);
            Assert.DoesNotContain(headers[0].Header!.Descendants<Text>(), t => t.Text == "MARKER-SECOND-HEADER");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void PartName_matches_the_real_file_name_case_insensitively()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var firstHeader = doc.MainDocumentPart!.HeaderParts.First();
            var upperCaseName = Path.GetFileName(firstHeader.Uri.OriginalString).ToUpperInvariant();

            var anchor = LocationResolver.Resolve(
                new Location { Type = LocationKind.DocumentEnd, Part = LayoutPart.Header, PartName = upperCaseName }, doc);

            // anchor.PartName reports the REAL (on-disk) casing, not the caller's requested casing.
            Assert.Equal(Path.GetFileName(firstHeader.Uri.OriginalString), anchor.PartName);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void AfterControl_on_Header_finds_a_control_that_lives_only_in_that_header()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);

            // Take the control from the part a partName-less header location actually resolves to (the
            // section default, header1.xml), not from HeaderParts.First(): this test is about AfterControl
            // INSIDE a header, so it must not smuggle in an assumption about which header that is.
            var targetHeader = doc.MainDocumentPart!.HeaderParts
                .Single(h => Path.GetFileName(h.Uri.OriginalString) == "header1.xml");
            var headerRoot = targetHeader.Header!;
            var controlId = headerRoot.Descendants<SdtElement>().Select(ReadId).First(id => id.HasValue)!.Value;

            var anchor = LocationResolver.Resolve(
                new Location { Type = LocationKind.AfterControl, ControlId = controlId, Part = LayoutPart.Header }, doc);
            Assert.Equal("header1.xml", anchor.PartName);

            var marker = new Run(new Text("MARKER-AFTER-HEADER-CONTROL"));
            anchor.InsertInline(marker);

            Assert.Contains(headerRoot.Descendants<Run>(), r => ReferenceEquals(r, marker));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void AfterControl_scoped_to_Header_does_not_find_a_control_that_only_exists_in_the_body()
    {
        // YourReference_Lbl (id -1130623254) lives in document.xml (the same real control
        // LocationResolverTests' body-level AfterControl tests already target) - proves resolution is
        // scoped to the CHOSEN part only, not a fallback search across the whole document.
        const int bodyOnlyControlId = -1130623254;
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var body = doc.MainDocumentPart!.Document!.Body!;
            Assert.Contains(body.Descendants<SdtElement>(), s => ReadId(s) == bodyOnlyControlId);

            var ex = Assert.Throws<NotFoundException>(() =>
                LocationResolver.Resolve(
                    new Location { Type = LocationKind.AfterControl, ControlId = bodyOnlyControlId, Part = LayoutPart.Header }, doc));

            Assert.Contains("No control with id", ex.Message);
            Assert.Contains(bodyOnlyControlId.ToString(), ex.Message);
            Assert.Equal(NotFoundTarget.Control, ex.TargetKind);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Unknown_PartName_throws_NotFoundException_naming_the_available_parts()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);

            var ex = Assert.Throws<NotFoundException>(() =>
                LocationResolver.Resolve(
                    new Location
                    {
                        Type = LocationKind.DocumentEnd,
                        Part = LayoutPart.Header,
                        PartName = "header-does-not-exist.xml",
                    },
                    doc));

            Assert.Contains("header-does-not-exist.xml", ex.Message);
            Assert.Contains("header1.xml", ex.Message);
            Assert.Equal(NotFoundTarget.NamedHeaderFooterPart, ex.TargetKind);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Layout_with_no_header_parts_throws_NotFoundException_when_Header_is_targeted()
    {
        // A synthetic layout (SyntheticLayout.Create) never adds any header/footer part at all.
        var path = SyntheticLayout.Create(SyntheticLayout.PlainParagraph("x"));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            Assert.Empty(doc.MainDocumentPart!.HeaderParts);

            var ex = Assert.Throws<NotFoundException>(() =>
                LocationResolver.Resolve(new Location { Type = LocationKind.DocumentEnd, Part = LayoutPart.Header }, doc));

            Assert.Contains("no header parts", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(NotFoundTarget.HeaderFooterParts, ex.TargetKind);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Layout_with_no_footer_parts_throws_NotFoundException_when_Footer_is_targeted()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.PlainParagraph("x"));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            Assert.Empty(doc.MainDocumentPart!.FooterParts);

            var ex = Assert.Throws<NotFoundException>(() =>
                LocationResolver.Resolve(new Location { Type = LocationKind.DocumentEnd, Part = LayoutPart.Footer }, doc));

            Assert.Contains("no footer parts", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(NotFoundTarget.HeaderFooterParts, ex.TargetKind);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Insert_field_into_a_corpus_header_reopens_clean_and_LayoutReader_reports_it_in_that_header_part()
    {
        // The grounded, primary-deliverable proof: a field control built exactly the way LayoutEditor.InsertField
        // builds one, inserted into a REAL corpus layout's header via Location.Part=Header, survives a full
        // save/reopen round trip clean on both OpenXmlValidator and LayoutValidator.Quick, and LayoutReader
        // correctly attributes it to that header part (not document.xml).
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            string expectedPartName;
            int controlId;
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var main = doc.MainDocumentPart!;
                // The SECTION DEFAULT header, named rather than taken from HeaderParts.First() (which is
                // header2.xml here) - see Partless_header_footer_target_resolves_to_the_first_sections_DEFAULT_part.
                expectedPartName = "header1.xml";

                var schema = SchemaProvider.FromLayout(doc);
                var field = SdtFactory.BuildField(schema, "/Header/CustomerAddress1", placeholderText: "HEADER-FIELD-ROUNDTRIP");
                controlId = field.SdtProperties!.GetFirstChild<SdtId>()!.Val!.Value;

                var anchor = LocationResolver.Resolve(new Location { Type = LocationKind.DocumentEnd, Part = LayoutPart.Header }, doc);
                Assert.Equal(expectedPartName, anchor.PartName);
                anchor.InsertInline(field);

                doc.MainDocumentPart!.Document!.Save();
                foreach (var header in main.HeaderParts)
                {
                    header.Header?.Save();
                }
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var openXmlErrors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened).ToList();
            Assert.Empty(openXmlErrors);

            var quick = LayoutValidator.Quick(reopened);
            Assert.Equal(0, quick.ErrorCount);

            var inventory = LayoutReader.Read(reopened);
            // The corpus's PRE-EXISTING /Header/CustomerAddress1 field (in document.xml) is untouched and
            // still there - a dataset path CAN legitimately be bound by more than one control (e.g. a
            // customer address shown in both a letterhead header and the body). What proves this edit
            // genuinely landed in the header, and not the body, is THIS SPECIFIC newly inserted control
            // (identified by its own w:id) being reported in the header part, and nowhere else.
            Assert.Contains(inventory.Controls, c =>
                c.SdtId == controlId &&
                c.Part == expectedPartName &&
                c.Kind == ControlKind.Field &&
                c.Alias == "#Nav: /Header/CustomerAddress1");
            Assert.DoesNotContain(inventory.Controls, c => c.SdtId == controlId && c.Part != expectedPartName);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- partName-less header/footer targeting resolves the SECTION DEFAULT, not the first part ----
    //
    // A Word document with distinct first-page/even-page headers carries three header parts whose
    // relationship order says nothing about which one is the everyday header. Measured across the corpus,
    // FIVE of the six layouts with header/footer parts had a different first part than first-section default,
    // and on three of them header1.xml is the EVEN-PAGE header - so a partName-less insert landed on even
    // pages only and was invisible on a one-page document. Footers went to the FIRST-PAGE footer, which is
    // why previewing page 1 looked correct and hid it. The cases below are the real per-file answers, so they
    // fail if the selection ever drifts back to positional.

    [Theory]
    // header1 = even, header2 = default, header3 = first; footers likewise (JobQuote is also the corpus's
    // only multi-section layout, 4 w:sectPr - the first one in document order is the one that governs).
    [InlineData(Corpus.JobQuote, LayoutPart.Header, "header2.xml")]
    [InlineData(Corpus.JobQuote, LayoutPart.Footer, "footer2.xml")]
    [InlineData(Corpus.StandardSalesQuote, LayoutPart.Header, "header2.xml")]
    [InlineData(Corpus.StandardSalesQuote, LayoutPart.Footer, "footer2.xml")]
    [InlineData(Corpus.StandardPurchaseOrder, LayoutPart.Header, "header2.xml")]
    [InlineData(Corpus.StandardPurchaseOrder, LayoutPart.Footer, "footer2.xml")]
    // Here the default IS part 1 - but its relationship order puts header2/footer2 first, so a positional
    // rule got these wrong in the opposite direction. Both directions must be pinned or the test proves
    // nothing about which rule is in force.
    [InlineData(Corpus.SalesInvoice, LayoutPart.Header, "header1.xml")]
    [InlineData(Corpus.SalesInvoice, LayoutPart.Footer, "footer1.xml")]
    public void Partless_header_footer_target_resolves_to_the_first_sections_DEFAULT_part(
        string corpusFile, LayoutPart part, string expected)
    {
        using var doc = WordprocessingDocument.Open(Corpus.Path(corpusFile), false);
        var (_, partName) = LocationResolver.ResolvePart(doc, part, null);
        Assert.Equal(expected, partName);
    }

    [Theory]
    [InlineData(Corpus.JobQuote, LayoutPart.Header, "header3.xml")]
    [InlineData(Corpus.JobQuote, LayoutPart.Footer, "footer1.xml")]
    public void An_explicit_partName_still_wins_over_the_section_default(
        string corpusFile, LayoutPart part, string requested)
    {
        // The default-part rule must not override a caller who named a part - targeting the first-page or
        // even-page header deliberately is a legitimate thing to want.
        using var doc = WordprocessingDocument.Open(Corpus.Path(corpusFile), false);
        var (_, partName) = LocationResolver.ResolvePart(doc, part, requested);
        Assert.Equal(requested, partName);
    }

}
