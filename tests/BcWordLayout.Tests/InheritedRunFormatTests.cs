using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

/// <summary>
/// Covers GitHub issue #122: inline inserts take the run formatting of the text before them
/// (<see cref="InheritedRunFormat"/>), so a separator, field, label or page number matches its neighbours
/// instead of rendering with the paragraph's and the style's defaults.
/// </summary>
public class InheritedRunFormatTests
{
    // StandardSalesQuote header2: every run of the title paragraph carries w:rStyle 'Strong', and so do both
    // controls' own w:sdtPr/w:rPr. The issue's headline case.
    private const int DocumentTitleLblId = 1061445566;
    private const int DocumentNoId = 1543556682;

    private static string CopyOfCorpus(string corpusFile)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-inheritfmt-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    private static Location AfterInHeader2(int controlId) => new()
    {
        Type = LocationKind.AfterControl,
        ControlId = controlId,
        Part = LayoutPart.Header,
        PartName = "header2.xml",
    };

    private static Header Header2(WordprocessingDocument doc) => doc.MainDocumentPart!.HeaderParts
        .Single(h => h.Uri.OriginalString.EndsWith("header2.xml", StringComparison.Ordinal)).Header!;

    private static void SaveAll(WordprocessingDocument doc)
    {
        doc.MainDocumentPart!.Document!.Save();
        foreach (var header in doc.MainDocumentPart.HeaderParts)
        {
            header.Header?.Save();
        }
    }

    private static int ErrorCount(string path)
    {
        using var doc = WordprocessingDocument.Open(path, false);
        return new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc).Count();
    }

    private static Run RunWithText(OpenXmlElement root, string text) =>
        root.Descendants<Run>().Single(r => r.GetFirstChild<Text>()?.Text == text);

    // ---- end to end, on the stock layouts the issue names ----

    [Fact]
    public void InsertText_after_a_bold_stock_control_takes_its_character_style()
    {
        var path = CopyOfCorpus(Corpus.StandardSalesQuote);
        try
        {
            var errorsBefore = ErrorCount(path);
            EditResult result;
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                result = LayoutEditor.InsertText(doc, " | ", AfterInHeader2(DocumentNoId));
                SaveAll(doc);
            }

            Assert.Contains("takes the formatting of the text before it (character style 'Strong')", result.Summary);

            using var reopened = WordprocessingDocument.Open(path, false);
            var run = RunWithText(Header2(reopened), " | ");
            Assert.Equal("Strong", run.RunProperties?.RunStyle?.Val?.Value);
            Assert.Equal(errorsBefore, new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened).Count());
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertField_after_a_bold_stock_control_styles_its_content_and_its_sdtPr()
    {
        var path = CopyOfCorpus(Corpus.StandardSalesQuote);
        try
        {
            int id;
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                id = LayoutEditor.InsertField(doc, "/Header/DocumentDate", AfterInHeader2(DocumentTitleLblId)).ControlId;
                SaveAll(doc);
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var control = Header2(reopened).Descendants<SdtRun>().Single(s => SdtInspector.ReadControlId(s) == id);

            // The control's own rPr comes first in its sdtPr, the corpus order.
            var sdtRPr = Assert.IsType<RunProperties>(control.SdtProperties!.FirstChild);
            Assert.Equal("Strong", sdtRPr.RunStyle?.Val?.Value);
            Assert.All(control.Descendants<Run>(), r => Assert.Equal("Strong", r.RunProperties?.RunStyle?.Val?.Value));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Bold_false_overrides_an_inherited_bold_character_style_with_an_explicit_off()
    {
        // Removing w:b cannot un-bold a run whose bold comes from its character style; w:b w:val="0" can.
        var path = CopyOfCorpus(Corpus.StandardSalesQuote);
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                LayoutEditor.InsertText(doc, " | ", AfterInHeader2(DocumentNoId), new TextFormat { Bold = false });
                SaveAll(doc);
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var rPr = RunWithText(Header2(reopened), " | ").RunProperties!;
            Assert.Equal("Strong", rPr.RunStyle?.Val?.Value);
            Assert.False(rPr.Bold?.Val?.Value ?? true);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Explicit_knobs_win_over_what_the_run_inherits()
    {
        var path = CopyOfCorpus(Corpus.StandardSalesQuote);
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                LayoutEditor.InsertText(doc, " | ", AfterInHeader2(DocumentNoId),
                    new TextFormat { Bold = true, FontSizePoints = 7 });
                SaveAll(doc);
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var rPr = RunWithText(Header2(reopened), " | ").RunProperties!;
            Assert.Equal("Strong", rPr.RunStyle?.Val?.Value); // still inherited
            Assert.NotNull(rPr.Bold);
            Assert.Equal("14", rPr.FontSize?.Val?.Value); // 7 pt = 14 half-points, the knob's value
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertPageNumber_into_an_empty_stock_cell_takes_the_paragraph_mark_formatting_on_every_run()
    {
        // JobQuote's table 2 has empty cells whose paragraph mark is Segoe UI 8 pt - with no text to follow,
        // Word formats typed text from the paragraph mark, and so does the insert.
        var path = CopyOfCorpus(Corpus.JobQuote);
        try
        {
            var location = new Location { Type = LocationKind.TableCell, TableIndex = 2, Row = 2, Col = 1 };
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                // The fixture's premise: the cell's last paragraph, where the insert lands, holds no text.
                var (cell, _) = LocationResolver.ResolveCellElement(location, doc);
                var last = cell.Elements<Paragraph>().Last();
                Assert.Equal("Segoe UI", last.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<RunFonts>()?.Ascii?.Value);
                Assert.Empty(last.Descendants<Run>());

                LayoutEditor.InsertPageNumber(doc, location);
                SaveAll(doc);
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var runs = reopened.MainDocumentPart!.Document!.Body!.Descendants<Paragraph>()
                .Single(p => p.Descendants<FieldCode>().Any(f => f.Text.Contains("NUMPAGES", StringComparison.Ordinal)))
                .Elements<Run>().ToList();
            Assert.Equal(11, runs.Count);
            Assert.All(runs, r =>
            {
                Assert.Equal("Segoe UI", r.RunProperties?.RunFonts?.Ascii?.Value);
                Assert.Equal("16", r.RunProperties?.FontSize?.Val?.Value);
            });

            // The cached results keep their corpus-shape noProof under the inherited formatting.
            Assert.Equal(2, runs.Count(r => r.RunProperties?.GetFirstChild<NoProof>() is not null));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void An_insert_that_opens_a_new_paragraph_stays_unformatted()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            EditResult result;
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                result = LayoutEditor.InsertText(doc, "Plain", new Location { Type = LocationKind.DocumentEnd });
                SaveAll(doc);
            }

            Assert.DoesNotContain("formatting of the text before it", result.Summary);
            using var reopened = WordprocessingDocument.Open(path, false);
            Assert.Null(RunWithText(reopened.MainDocumentPart!.Document!.Body!, "Plain").RunProperties);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_placeholder_showing_neighbour_passes_on_its_formatting_but_not_the_placeholder_style()
    {
        var path = CopyOfCorpus(Corpus.StandardSalesQuote);
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                // Turn DocumentNo into a control showing placeholder text: no sdtPr/w:rPr of its own, and a
                // content run in Word's PlaceholderText style.
                var control = Header2(doc).Descendants<SdtRun>().Single(s => SdtInspector.ReadControlId(s) == DocumentNoId);
                control.SdtProperties!.RemoveAllChildren<RunProperties>();
                foreach (var run in control.Descendants<Run>())
                {
                    run.RunProperties = new RunProperties(new RunStyle { Val = "PlaceholderText" }, new Italic());
                }

                LayoutEditor.InsertText(doc, " | ", AfterInHeader2(DocumentNoId));
                SaveAll(doc);
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var rPr = RunWithText(Header2(reopened), " | ").RunProperties!;
            Assert.Null(rPr.RunStyle);
            Assert.NotNull(rPr.Italic);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_picture_control_anchor_passes_on_no_formatting()
    {
        // A drawing run's rPr is not text formatting - Word even writes w:w="0" there (GitHub issue #54).
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var pictureId = LayoutEditor.InsertPicture(
                    doc, "/Header/CompanyPicture", new Location { Type = LocationKind.DocumentEnd }).ControlId;
                var picture = doc.MainDocumentPart!.Document!.Body!.Descendants<SdtRun>()
                    .Single(s => SdtInspector.ReadControlId(s) == pictureId);
                picture.Descendants<Run>().Single().RunProperties!.AppendChild(new CharacterScale { Val = 0 });

                LayoutEditor.InsertText(doc, "Caption",
                    new Location { Type = LocationKind.AfterControl, ControlId = pictureId });
                SaveAll(doc);
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            Assert.Null(RunWithText(reopened.MainDocumentPart!.Document!.Body!, "Caption").RunProperties);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- the rule itself, on synthetic paragraphs ----

    [Fact]
    public void Before_takes_the_last_text_run_and_an_unformatted_one_passes_on_nothing()
    {
        var bold = new Run(new RunProperties(new Bold()), new Text("Total"));
        var plain = new Run(new Text(":"));

        Assert.Null(InheritedRunFormat.Before(new Paragraph(bold, plain), insertAfter: null));
        Assert.NotNull(InheritedRunFormat.Before(new Paragraph(plain.CloneNode(true), bold.CloneNode(true)), null)?.Bold);
    }

    [Fact]
    public void Before_steps_back_over_pictures_and_bookmarks_to_the_text()
    {
        var italic = new Run(new RunProperties(new Italic()), new Text("Logo:"));
        var drawing = new Run(new RunProperties(new CharacterScale { Val = 0 }), new Drawing());
        var paragraph = new Paragraph(italic, new BookmarkStart { Id = "1", Name = "b" }, drawing, new BookmarkEnd { Id = "1" });

        var rPr = InheritedRunFormat.Before(paragraph, insertAfter: null);

        Assert.NotNull(rPr?.Italic);
        Assert.Null(rPr!.GetFirstChild<CharacterScale>());
    }

    [Fact]
    public void Before_reads_an_empty_paragraphs_mark_without_its_revision_marks_or_hiding()
    {
        var paragraph = new Paragraph(new ParagraphProperties(new ParagraphMarkRunProperties(
            new Inserted { Id = "1", Author = "x" },
            new Bold(),
            new Vanish(),
            new FontSize { Val = "16" })));

        var rPr = InheritedRunFormat.Before(paragraph, insertAfter: null);

        Assert.NotNull(rPr);
        Assert.NotNull(rPr!.Bold);
        Assert.Equal("16", rPr.FontSize?.Val?.Value);
        Assert.Null(rPr.Vanish);
        Assert.Empty(rPr.Elements<Inserted>());
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(new Run(rPr, new Text("x"))));
    }

    [Fact]
    public void Before_returns_a_detached_copy_each_time()
    {
        var source = new Run(new RunProperties(new Bold()), new Text("a"));
        var paragraph = new Paragraph(source);

        var first = InheritedRunFormat.Before(paragraph, null)!;
        first.Italic = new Italic();

        Assert.Null(first.Parent);
        Assert.Null(source.RunProperties!.Italic);
        Assert.Null(InheritedRunFormat.Before(paragraph, null)!.Italic);
    }
}
