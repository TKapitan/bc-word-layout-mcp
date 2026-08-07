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
/// Covers GitHub issue #29: emitting the Word <c>PAGE</c>/<c>NUMPAGES</c> field-code idiom every stock BC
/// document header carries — <see cref="PageNumberFieldFactory"/>'s exact corpus shape (instruction
/// spacing, <c>w:noProof</c> on the cached run, the literal <c>" / "</c> separator),
/// <see cref="LayoutEditor.InsertPageNumber"/>'s placement/round-trip, and the <c>insert_page_number</c>
/// MCP tool. The reference shape was extracted from the four corpus captures
/// (StandardSalesQuote header2/3, StandardPurchaseOrder header2, StandardSalesInvoiceVatSpec header1,
/// SalespersonCommission header1), which all agree on it.
/// </summary>
public class PageNumberFieldTests
{
    private static string CopyOfCorpus(string corpusFile)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-pagenum-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    // ---- PageNumberFieldFactory: the exact corpus shape ----

    /// <summary>Asserts one five-run field construct starting at <paramref name="runs"/>[<paramref name="start"/>].</summary>
    private static void AssertFieldConstruct(IReadOnlyList<Run> runs, int start, string expectedInstruction)
    {
        Assert.Equal(FieldCharValues.Begin, runs[start].GetFirstChild<FieldChar>()?.FieldCharType?.Value);

        var instruction = runs[start + 1].GetFirstChild<FieldCode>();
        Assert.NotNull(instruction);
        Assert.Equal(expectedInstruction, instruction!.Text);
        Assert.Equal(SpaceProcessingModeValues.Preserve, instruction.Space?.Value);

        Assert.Equal(FieldCharValues.Separate, runs[start + 2].GetFirstChild<FieldChar>()?.FieldCharType?.Value);

        // The cached result: "1" under w:noProof — present in every corpus capture.
        var cached = runs[start + 3];
        Assert.Equal("1", cached.GetFirstChild<Text>()?.Text);
        Assert.NotNull(cached.RunProperties?.GetFirstChild<NoProof>());

        Assert.Equal(FieldCharValues.End, runs[start + 4].GetFirstChild<FieldChar>()?.FieldCharType?.Value);
    }

    [Fact]
    public void BuildPageOfTotal_is_the_corpus_shape_PAGE_separator_NUMPAGES()
    {
        var runs = PageNumberFieldFactory.BuildPageOfTotal();

        Assert.Equal(11, runs.Count);
        AssertFieldConstruct(runs, 0, " PAGE  \\* Arabic  \\* MERGEFORMAT ");

        // The literal " / " separator run between the fields, whitespace preserved.
        var separator = runs[5].GetFirstChild<Text>();
        Assert.NotNull(separator);
        Assert.Equal(" / ", separator!.Text);
        Assert.Equal(SpaceProcessingModeValues.Preserve, separator.Space?.Value);

        AssertFieldConstruct(runs, 6, " NUMPAGES  \\* Arabic  \\* MERGEFORMAT ");
    }

    [Fact]
    public void BuildPageNumber_is_the_bare_five_run_PAGE_construct()
    {
        var runs = PageNumberFieldFactory.BuildPageNumber();

        Assert.Equal(5, runs.Count);
        AssertFieldConstruct(runs, 0, " PAGE  \\* Arabic  \\* MERGEFORMAT ");
    }

    // ---- LayoutEditor.InsertPageNumber: placement and round-trip ----

    [Fact]
    public void InsertPageNumber_into_a_header_round_trips_and_introduces_no_validation_errors()
    {
        var path = CopyOfCorpus(Corpus.JobQuote);
        try
        {
            EditResult result;
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                result = LayoutEditor.InsertPageNumber(doc, new Location
                {
                    Type = LocationKind.DocumentEnd,
                    Part = LayoutPart.Header,
                });

                Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc));

                doc.MainDocumentPart!.Document!.Save();
                foreach (var header in doc.MainDocumentPart.HeaderParts)
                {
                    header.Header?.Save();
                }
            }

            Assert.Equal("insert_page_number", result.Operation);
            Assert.Equal(0, result.ControlId);
            Assert.Equal("PageNumberField", result.Kind);
            Assert.Equal("header2.xml", result.Part); // the first section's DEFAULT header, not header1.xml

            // Reopen read-only: the field construct must be on disk, contiguous, in the targeted part.
            using var reopened = WordprocessingDocument.Open(path, false);
            var header2 = reopened.MainDocumentPart!.HeaderParts
                .Single(h => h.Uri.OriginalString.EndsWith("header2.xml", StringComparison.Ordinal));
            var instructions = header2.Header!.Descendants<FieldCode>().Select(f => f.Text).ToList();
            Assert.Contains(" PAGE  \\* Arabic  \\* MERGEFORMAT ", instructions);
            Assert.Contains(" NUMPAGES  \\* Arabic  \\* MERGEFORMAT ", instructions);

            // The eleven runs are consecutive siblings of one paragraph: begin..end, " / ", begin..end.
            var beginRun = header2.Header.Descendants<Run>()
                .First(r => r.GetFirstChild<FieldChar>()?.FieldCharType?.Value == FieldCharValues.Begin);
            var sequence = new List<Run> { beginRun };
            var next = beginRun.NextSibling();
            while (next is Run run && sequence.Count < 11)
            {
                sequence.Add(run);
                next = run.NextSibling();
            }

            Assert.Equal(11, sequence.Count);
            Assert.Equal(" / ", sequence[5].GetFirstChild<Text>()?.Text);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertPageNumber_without_total_emits_the_PAGE_field_only()
    {
        var path = CopyOfCorpus(Corpus.JobQuote);
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                LayoutEditor.InsertPageNumber(
                    doc,
                    new Location { Type = LocationKind.DocumentEnd, Part = LayoutPart.Footer },
                    includeTotal: false);
                doc.MainDocumentPart!.Document!.Save();
                foreach (var footer in doc.MainDocumentPart.FooterParts)
                {
                    footer.Footer?.Save();
                }
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var footer2 = reopened.MainDocumentPart!.FooterParts
                .Single(f => f.Uri.OriginalString.EndsWith("footer2.xml", StringComparison.Ordinal));
            var instructions = footer2.Footer!.Descendants<FieldCode>().Select(f => f.Text).ToList();
            Assert.Contains(" PAGE  \\* Arabic  \\* MERGEFORMAT ", instructions);
            Assert.DoesNotContain(" NUMPAGES  \\* Arabic  \\* MERGEFORMAT ", instructions);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertPageNumber_after_a_control_lands_in_the_same_paragraph_run_flow()
    {
        // The composed stock idiom: label, separator, fields — all inline in ONE paragraph. Anchoring
        // afterControl on an inline control must put the field runs into that control's own paragraph.
        var path = CopyOfCorpus(Corpus.JobQuote);
        try
        {
            int labelId;
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var label = LayoutEditor.InsertField(doc, "/Job/BillToAddress1", new Location
                {
                    Type = LocationKind.DocumentEnd,
                    Part = LayoutPart.Header,
                });
                labelId = label.ControlId;

                var result = LayoutEditor.InsertPageNumber(doc, new Location
                {
                    Type = LocationKind.AfterControl,
                    Part = LayoutPart.Header,
                    ControlId = labelId,
                });
                Assert.Equal("header2.xml", result.Part);

                doc.MainDocumentPart!.Document!.Save();
                foreach (var header in doc.MainDocumentPart.HeaderParts)
                {
                    header.Header?.Save();
                }
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var header2 = reopened.MainDocumentPart!.HeaderParts
                .Single(h => h.Uri.OriginalString.EndsWith("header2.xml", StringComparison.Ordinal));
            var anchor = header2.Header!.Descendants<SdtElement>()
                .Single(s => SdtInspector.ReadControlId(s) == labelId);
            var paragraph = Assert.IsType<Paragraph>(anchor.Parent);
            Assert.Equal(2, paragraph.Descendants<FieldCode>().Count());
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertPageNumber_applies_bold_and_size_to_every_run_and_keeps_noProof()
    {
        var path = CopyOfCorpus(Corpus.JobQuote);
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                LayoutEditor.InsertPageNumber(
                    doc,
                    new Location { Type = LocationKind.DocumentEnd, Part = LayoutPart.Header },
                    includeTotal: true,
                    new CellTextFormat { Bold = true, FontSizePoints = 8 });
                doc.MainDocumentPart!.Document!.Save();
                foreach (var header in doc.MainDocumentPart.HeaderParts)
                {
                    header.Header?.Save();
                }
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var header2 = reopened.MainDocumentPart!.HeaderParts
                .Single(h => h.Uri.OriginalString.EndsWith("header2.xml", StringComparison.Ordinal));
            var beginRun = header2.Header!.Descendants<Run>()
                .First(r => r.GetFirstChild<FieldChar>()?.FieldCharType?.Value == FieldCharValues.Begin);

            var sequence = new List<Run> { beginRun };
            var next = beginRun.NextSibling();
            while (next is Run run && sequence.Count < 11)
            {
                sequence.Add(run);
                next = run.NextSibling();
            }

            Assert.All(sequence, r =>
            {
                Assert.NotNull(r.RunProperties?.Bold);
                Assert.Equal("16", r.RunProperties!.FontSize?.Val?.Value); // 8 pt = 16 half-points
            });

            // The cached-result run keeps its corpus-shape noProof alongside the added styling.
            Assert.NotNull(sequence[3].RunProperties?.GetFirstChild<NoProof>());
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertPageNumber_rejects_an_alignment_like_every_other_inline_insert()
    {
        var path = CopyOfCorpus(Corpus.JobQuote);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var ex = Assert.Throws<ArgumentException>(() => LayoutEditor.InsertPageNumber(
                doc,
                new Location { Type = LocationKind.DocumentEnd, Part = LayoutPart.Header },
                includeTotal: true,
                new CellTextFormat { Alignment = "right" }));
            Assert.Contains("alignment", ex.Message, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- the insert_page_number MCP tool ----

    [Fact]
    public void InsertPageNumber_tool_returns_ok_with_post_edit_validation_and_scaffolds_a_missing_footer()
    {
        // A blank from-scratch build has no footer part at all: the tool must scaffold one (like
        // insert_field does) rather than dead-end, and the response must carry the usual envelope shape.
        var path = SyntheticLayout.Create(SyntheticLayout.PlainParagraph("body"));
        try
        {
            var response = EditTools.InsertPageNumber(path, "documentEnd", layoutPart: "footer");

            Assert.True(response.Ok, response.Error?.Message);
            var dto = Assert.IsType<EditResultDto>(response.Data);
            Assert.Equal("insert_page_number", dto.Operation);
            Assert.Equal(0, dto.ControlId);
            Assert.Equal("footer1.xml", dto.Part);
            Assert.Contains("was created and wired into the page setup", dto.Summary, StringComparison.Ordinal);
            Assert.NotNull(dto.QuickValidation);

            using var reopened = WordprocessingDocument.Open(path, false);
            var footer = Assert.Single(reopened.MainDocumentPart!.FooterParts);
            Assert.Contains(footer.Footer!.Descendants<FieldCode>(), f => f.Text.Contains("PAGE", StringComparison.Ordinal));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertPageNumber_tool_reports_invalid_argument_for_a_bad_location_type()
    {
        var response = EditTools.InsertPageNumber(Corpus.Path(Corpus.JobQuote), "nowhere");

        Assert.False(response.Ok);
        Assert.Equal("invalid_argument", response.Error!.Code);
        Assert.False(string.IsNullOrWhiteSpace(response.Error.Hint));
    }
}
