using BcWordLayout.Domain;
using BcWordLayout.McpHost;
using BcWordLayout.McpHost.Tools;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

/// <summary>
/// Covers <c>set_text_format</c> (<see cref="TextFormatEditor"/>, GitHub issue #15): restyling text that is
/// already in the layout, by control, cell, row or exact text, and the refusals that keep it from guessing.
/// </summary>
public class TextFormatEditorTests
{
    // StandardSalesQuote header2: DocumentNo carries 'Strong' on its sdtPr/rPr and its runs.
    private const int DocumentNoId = 1543556682;

    private static string CopyOfCorpus(string corpusFile)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-textformat-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    private static int SchemaErrors(string path)
    {
        using var doc = WordprocessingDocument.Open(path, false);
        return new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc).Count();
    }

    private static EditResultDto Ok(ToolResponse response)
    {
        Assert.True(response.Ok, response.Error?.Message);
        return Assert.IsType<EditResultDto>(response.Data);
    }

    private static void Refused(ToolResponse response, string code, string fragment)
    {
        Assert.False(response.Ok);
        Assert.Equal(code, response.Error!.Code);
        Assert.Contains(fragment, response.Error.Message, StringComparison.Ordinal);
    }

    /// <summary>A fresh plain table at the end of the body, returning its index.</summary>
    private static int NewPlainTable(string path, int rows = 2, int columns = 3)
    {
        var response = TableTools.InsertTable(path, rows, TableArgs.Columns(columns), locationType: "documentEnd");
        Assert.True(response.Ok, response.Error?.Message);
        return Assert.IsType<TableEditResultDto>(response.Data).TableIndex;
    }

    private static TableRow BodyRow(WordprocessingDocument doc, int tableIndex, int row) =>
        TableGridNavigator.Tables(doc.MainDocumentPart!.Document!.Body!)[tableIndex].Elements<TableRow>().ElementAt(row);

    private static SdtElement Control(WordprocessingDocument doc, int id) =>
        PartRoots(doc).SelectMany(r => r.Descendants<SdtElement>()).Single(s => SdtInspector.ReadControlId(s) == id);

    private static IEnumerable<OpenXmlElement> PartRoots(WordprocessingDocument doc) =>
        new OpenXmlElement[] { doc.MainDocumentPart!.Document! }
            .Concat(doc.MainDocumentPart.HeaderParts.Select(h => (OpenXmlElement)h.Header!))
            .Concat(doc.MainDocumentPart.FooterParts.Select(f => (OpenXmlElement)f.Footer!));

    // ---- control ----

    [Fact]
    public void A_stock_control_takes_the_format_on_its_own_rPr_and_its_text_over_what_it_has()
    {
        var path = CopyOfCorpus(Corpus.StandardSalesQuote);
        try
        {
            var errorsBefore = SchemaErrors(path);

            var dto = Ok(EditTools.SetTextFormat(path, "control", controlId: DocumentNoId,
                format: new TextFormatArgument { Color = "0070C0", Italic = true }));

            Assert.Equal("set_text_format", dto.Operation);
            Assert.Equal(DocumentNoId, dto.ControlId);
            Assert.Equal("header2.xml", dto.Part);
            Assert.Contains("italic, colour 0070C0", dto.Summary, StringComparison.Ordinal);

            using var reopened = WordprocessingDocument.Open(path, false);
            var control = Control(reopened, DocumentNoId);
            var own = Assert.IsType<RunProperties>(control.SdtProperties!.FirstChild);
            foreach (var rPr in control.Descendants<RunProperties>().Prepend(own))
            {
                Assert.Equal("Strong", rPr.RunStyle?.Val?.Value); // kept: only the members passed change
                Assert.Equal("0070C0", rPr.Color?.Val?.Value);
                Assert.NotNull(rPr.Italic);
            }

            Assert.Equal(errorsBefore, SchemaErrors(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Alignment_on_an_inline_control_aligns_the_line_it_sits_in()
    {
        var path = CopyOfCorpus(Corpus.StandardSalesQuote);
        try
        {
            Ok(EditTools.SetTextFormat(path, "control", controlId: DocumentNoId, alignment: "right"));

            using var reopened = WordprocessingDocument.Open(path, false);
            var line = Control(reopened, DocumentNoId).Ancestors<Paragraph>().First();
            Assert.Equal(JustificationValues.Right, line.ParagraphProperties?.Justification?.Val?.Value);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_cell_level_control_is_formatted_with_its_whole_cell()
    {
        // The BC header address fields wrap a whole w:tc; formatting the control formats the cell's text.
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            int id;
            using (var probe = WordprocessingDocument.Open(path, false))
            {
                id = PartRoots(probe).SelectMany(r => r.Descendants<SdtCell>())
                    .First(c => SdtInspector.Classify(c.SdtProperties) == SdtInspector.Classification.Bound
                                && c.Descendants<Text>().Any())
                    .SdtProperties!.GetFirstChild<SdtId>()!.Val!.Value;
            }

            Ok(EditTools.SetTextFormat(path, "control", controlId: id, format: new TextFormatArgument { FontSizePoints = 7 }));

            using var reopened = WordprocessingDocument.Open(path, false);
            var control = Control(reopened, id);
            Assert.Equal("14", control.SdtProperties!.GetFirstChild<RunProperties>()?.FontSize?.Val?.Value);
            Assert.All(control.Descendants<Run>().Where(r => r.GetFirstChild<Text>() is not null),
                r => Assert.Equal("14", r.RunProperties?.FontSize?.Val?.Value));
            Assert.All(control.Descendants<Paragraph>(),
                p => Assert.Equal("14", p.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<FontSize>()?.Val?.Value));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_picture_control_is_refused_it_has_no_text()
    {
        var path = CopyOfCorpus(Corpus.StandardSalesQuote);
        try
        {
            int id;
            using (var probe = WordprocessingDocument.Open(path, false))
            {
                id = PartRoots(probe).SelectMany(r => r.Descendants<SdtElement>())
                    .First(s => SdtInspector.Classify(s.SdtProperties) == SdtInspector.Classification.Picture)
                    .SdtProperties!.GetFirstChild<SdtId>()!.Val!.Value;
            }

            Refused(EditTools.SetTextFormat(path, "control", controlId: id, format: new TextFormatArgument { Bold = true }),
                "invalid_argument", "is a picture control: it has no text to format");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void An_unknown_control_id_is_not_found()
    {
        var path = CopyOfCorpus(Corpus.StandardSalesQuote);
        try
        {
            Refused(EditTools.SetTextFormat(path, "control", controlId: 424242, format: new TextFormatArgument { Bold = true }),
                "not_found", "No control with id 424242");
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- cell and row ----

    [Fact]
    public void A_whole_row_is_formatted_and_aligned_in_one_call()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var table = NewPlainTable(path);
            for (var col = 0; col < 3; col++)
            {
                Assert.True(EditTools.SetCellText(path, table, 0, col, $"Caption {col}").Ok);
            }

            var dto = Ok(EditTools.SetTextFormat(path, "row", tableIndex: table, row: 0,
                format: new TextFormatArgument { Bold = true, FontSizePoints = 8 }, alignment: "center"));
            Assert.Equal(table, dto.TableIndex);
            Assert.Contains("all 3 cell(s) of table", dto.Summary, StringComparison.Ordinal);

            using var reopened = WordprocessingDocument.Open(path, false);
            foreach (var cell in BodyRow(reopened, table, 0).Elements<TableCell>())
            {
                var paragraph = cell.GetFirstChild<Paragraph>()!;
                Assert.Equal(JustificationValues.Center, paragraph.ParagraphProperties?.Justification?.Val?.Value);
                var rPr = paragraph.GetFirstChild<Run>()!.RunProperties!;
                Assert.NotNull(rPr.Bold);
                Assert.Equal("16", rPr.FontSize?.Val?.Value);
            }

            // The other row is untouched.
            Assert.All(BodyRow(reopened, table, 1).Descendants<ParagraphProperties>(),
                p => Assert.Null(p.ParagraphMarkRunProperties));
            var newTable = TableGridNavigator.Tables(reopened.MainDocumentPart!.Document!.Body!)[table];
            Assert.DoesNotContain(
                new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened.MainDocumentPart!),
                e => e.Node?.Ancestors<Table>().Contains(newTable) == true);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void An_empty_cell_keeps_the_format_on_its_mark_so_a_later_insert_takes_it()
    {
        // Format first, fill later: the paragraph mark carries the formatting, and an insert into an empty
        // paragraph takes the mark's formatting (GitHub issue #122).
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var table = NewPlainTable(path);

            var dto = Ok(EditTools.SetTextFormat(path, "cell", tableIndex: table, row: 0, col: 1,
                format: new TextFormatArgument { Bold = true, FontName = "Segoe UI" }));
            Assert.Contains("holds no text yet", dto.Summary, StringComparison.Ordinal);

            Assert.True(EditTools.InsertText(path, "Later", "tableCell", tableIndex: table, row: 0, col: 1).Ok);

            using var reopened = WordprocessingDocument.Open(path, false);
            var run = BodyRow(reopened, table, 0).Elements<TableCell>().ElementAt(1).Descendants<Run>().Single();
            Assert.NotNull(run.RunProperties?.Bold);
            Assert.Equal("Segoe UI", run.RunProperties?.RunFonts?.Ascii?.Value);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_cell_formats_the_field_inside_it_on_the_fields_own_rPr_too()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var table = NewPlainTable(path);
            var field = Assert.IsType<EditResultDto>(
                EditTools.InsertField(path, "/Header/ExternalDocumentNo", "tableCell", tableIndex: table, row: 1, col: 0).Data);

            Ok(EditTools.SetTextFormat(path, "cell", tableIndex: table, row: 1, col: 0,
                format: new TextFormatArgument { Underline = true }));

            using var reopened = WordprocessingDocument.Open(path, false);
            var control = Control(reopened, field.ControlId);
            Assert.Equal(UnderlineValues.Single, control.SdtProperties!.GetFirstChild<RunProperties>()?.Underline?.Val?.Value);
            Assert.All(control.Descendants<Run>(), r => Assert.Equal(UnderlineValues.Single, r.RunProperties?.Underline?.Val?.Value));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_nested_table_in_the_cell_is_left_alone()
    {
        // StandardPurchaseOrder's default footer: table 0's outer cell holds a nested 1x1 table (GitHub issue
        // #86). The nested table is a table of its own to address; formatting the outer cell must not reach it.
        var path = CopyOfCorpus(Corpus.StandardPurchaseOrder);
        try
        {
            Ok(EditTools.SetTextFormat(path, "cell", tableIndex: 0, row: 0, col: 0, layoutPart: "footer",
                format: new TextFormatArgument { Italic = true }));

            using var reopened = WordprocessingDocument.Open(path, false);
            var footer = reopened.MainDocumentPart!.FooterParts
                .Single(f => f.Uri.OriginalString.EndsWith("footer2.xml", StringComparison.Ordinal)).Footer!;
            var outer = footer.Elements<Table>().First().Descendants<TableCell>().First();
            var nested = outer.Elements<Table>().Single();
            Assert.DoesNotContain(nested.Descendants<RunProperties>(), r => r.Italic is not null);
            Assert.DoesNotContain(nested.Descendants<ParagraphMarkRunProperties>(), m => m.GetFirstChild<Italic>() is not null);
            Assert.All(outer.Elements<Paragraph>(),
                p => Assert.NotNull(p.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<Italic>()));
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- text ----

    [Fact]
    public void Text_inside_a_run_is_split_out_and_only_it_is_formatted()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            Assert.True(EditTools.InsertText(path, "Sales Zebra Confirmation", "documentEnd",
                format: new TextFormatArgument { FontSizePoints = 18 }).Ok);

            var dto = Ok(EditTools.SetTextFormat(path, "text", searchText: "Zebra",
                format: new TextFormatArgument { Italic = true }));
            Assert.Contains("split out of the run", dto.Summary, StringComparison.Ordinal);

            using var reopened = WordprocessingDocument.Open(path, false);
            var runs = reopened.MainDocumentPart!.Document!.Body!.Descendants<Paragraph>().Last().Elements<Run>().ToList();
            Assert.Equal(["Sales ", "Zebra", " Confirmation"], runs.Select(r => r.InnerText));
            Assert.All(runs, r => Assert.Equal("36", r.RunProperties?.FontSize?.Val?.Value)); // all keep the run's size
            Assert.Null(runs[0].RunProperties!.Italic);
            Assert.NotNull(runs[1].RunProperties!.Italic);
            Assert.Null(runs[2].RunProperties!.Italic);
            Assert.All(runs.SelectMany(r => r.Elements<Text>()), t => Assert.Equal(SpaceProcessingModeValues.Preserve, t.Space?.Value));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Text_that_is_a_whole_run_is_formatted_in_place()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            Assert.True(EditTools.InsertText(path, "ORDER LINES", "documentEnd").Ok);

            var dto = Ok(EditTools.SetTextFormat(path, "text", searchText: "ORDER LINES",
                format: new TextFormatArgument { Bold = true }, alignment: "center"));
            Assert.DoesNotContain("split", dto.Summary, StringComparison.Ordinal);

            using var reopened = WordprocessingDocument.Open(path, false);
            var paragraph = reopened.MainDocumentPart!.Document!.Body!.Descendants<Paragraph>().Last();
            Assert.NotNull(paragraph.Elements<Run>().Single().RunProperties?.Bold);
            Assert.Equal(JustificationValues.Center, paragraph.ParagraphProperties?.Justification?.Val?.Value);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Text_that_occurs_more_than_once_is_refused_with_the_count()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            Assert.True(EditTools.InsertText(path, "Twice", "documentEnd").Ok);
            Assert.True(EditTools.InsertText(path, "Twice", "documentEnd").Ok);
            var before = File.ReadAllBytes(path);

            Refused(EditTools.SetTextFormat(path, "text", searchText: "Twice", format: new TextFormatArgument { Bold = true }),
                "invalid_argument", "'Twice' occurs 2 times in the document body");
            Assert.Equal(before, File.ReadAllBytes(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Text_inside_a_bound_control_is_refused_naming_the_control_to_target_instead()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var table = NewPlainTable(path);
            var field = Assert.IsType<EditResultDto>(
                EditTools.InsertField(path, "/Header/ExternalDocumentNo", "tableCell", tableIndex: table, row: 0, col: 0).Data);
            string placeholder;
            using (var probe = WordprocessingDocument.Open(path, false))
            {
                placeholder = Control(probe, field.ControlId).Descendants<Text>().First().Text;
            }

            // Only meaningful if the placeholder text is unique; it carries the field's own name.
            Refused(EditTools.SetTextFormat(path, "text", searchText: placeholder, format: new TextFormatArgument { Bold = true }),
                "invalid_argument", $"target 'control', controlId {field.ControlId}");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Text_that_is_not_there_is_not_found()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            Refused(EditTools.SetTextFormat(path, "text", searchText: "no such text anywhere", format: new TextFormatArgument { Bold = true }),
                "not_found", "No text containing 'no such text anywhere'");
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- refusals before anything is touched ----

    [Theory]
    [InlineData("control", null, 0, null, null, "target 'control' needs controlId and does not use tableIndex")]
    [InlineData("row", null, 0, 0, 1, "target 'row' does not use col (a row target formats every cell")]
    [InlineData("cell", null, 0, null, null, "target 'cell' needs row, col")]
    [InlineData("text", 5, null, null, null, "target 'text' needs searchText and does not use controlId")]
    public void A_target_missing_its_arguments_or_carrying_others_is_refused(
        string target, int? controlId, int? tableIndex, int? row, int? col, string fragment)
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var before = File.ReadAllBytes(path);
            Refused(EditTools.SetTextFormat(path, target, controlId: controlId, tableIndex: tableIndex, row: row, col: col,
                format: new TextFormatArgument { Bold = true }), "invalid_argument", fragment);
            Assert.Equal(before, File.ReadAllBytes(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_control_target_refuses_a_part_since_ids_are_found_in_every_part()
    {
        var path = CopyOfCorpus(Corpus.StandardSalesQuote);
        try
        {
            Refused(EditTools.SetTextFormat(path, "control", controlId: DocumentNoId, layoutPart: "header",
                format: new TextFormatArgument { Bold = true }), "invalid_argument", "a control is found by its id in any part");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_call_with_nothing_to_change_is_refused()
    {
        var path = CopyOfCorpus(Corpus.StandardSalesQuote);
        try
        {
            Refused(EditTools.SetTextFormat(path, "control", controlId: DocumentNoId, format: new TextFormatArgument()),
                "invalid_argument", "set_text_format has nothing to change");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void An_unknown_target_is_refused_naming_the_valid_ones()
    {
        var path = CopyOfCorpus(Corpus.StandardSalesQuote);
        try
        {
            Refused(EditTools.SetTextFormat(path, "paragraph", format: new TextFormatArgument { Bold = true }),
                "invalid_argument", "Use 'control', 'cell', 'row', or 'text'.");
        }
        finally
        {
            File.Delete(path);
        }
    }
}
