using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

/// <summary>
/// Covers <see cref="LayoutEditor"/> directly (no MCP tool layer): each test opens a temp COPY of a
/// corpus/synthetic layout editable, applies one operation, saves, then reopens read-only to assert the
/// on-disk result — mirroring the round-trip style already used by <c>SdtFactoryTests</c> and
/// <c>LocationResolverTests</c>.
/// </summary>
public class LayoutEditorTests
{
    // Real, verified paths from tests/corpus/SalesInvoiceForSubscriptionBilling.docx (Standard_Sales_Invoice/1306) —
    // same ones SdtFactoryTests/LocationResolverTests already rely on.
    private const string FieldPath = "/Header/CustomerAddress1";
    private const string LabelPath = "/Header/Contact_Lbl";
    private const int CellLevelLabelId = -1130623254; // YourReference_Lbl, a cell-level SdtCell control.

    private static int? ReadId(SdtElement sdt) =>
        sdt.GetFirstChild<SdtProperties>()?.GetFirstChild<SdtId>()?.Val?.Value;

    private static string CopyOfCorpus(string corpusFile)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-layouteditor-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    // ---- insert_field ----

    [Fact]
    public void InsertField_at_documentEnd_is_readable_back_as_a_Field_with_the_expected_xpath_and_passes_validation()
    {
        var path = Corpus.UnlinkedCopy(Corpus.SalesInvoice);
        try
        {
            EditResult result;
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                result = LayoutEditor.InsertField(doc, FieldPath, new Location { Type = LocationKind.DocumentEnd });
                doc.MainDocumentPart!.Document!.Save();
            }

            Assert.Equal("insert_field", result.Operation);
            Assert.Equal("Field", result.Kind);
            Assert.Equal("document.xml", result.Part);
            Assert.Equal("/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:CustomerAddress1[1]", result.XPath);
            Assert.Equal("#Nav: /Header/CustomerAddress1", result.Alias);

            using var reopened = WordprocessingDocument.Open(path, false);
            var openXmlErrors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened).ToList();
            Assert.Empty(openXmlErrors);

            var quick = LayoutValidator.Quick(reopened);
            Assert.Equal(0, quick.ErrorCount);

            var inventory = LayoutReader.Read(reopened);
            Assert.Contains(inventory.Controls, c =>
                c.SdtId == result.ControlId &&
                c.Kind == ControlKind.Field &&
                c.XPath == result.XPath);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertLabel_at_documentEnd_is_readable_back_as_a_Label_and_passes_validation()
    {
        var path = Corpus.UnlinkedCopy(Corpus.SalesInvoice);
        try
        {
            EditResult result;
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                result = LayoutEditor.InsertLabel(doc, LabelPath, new Location { Type = LocationKind.DocumentEnd });
                doc.MainDocumentPart!.Document!.Save();
            }

            Assert.Equal("insert_label", result.Operation);
            Assert.Equal("Label", result.Kind);

            using var reopened = WordprocessingDocument.Open(path, false);
            var openXmlErrors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened).ToList();
            Assert.Empty(openXmlErrors);

            var quick = LayoutValidator.Quick(reopened);
            Assert.Equal(0, quick.ErrorCount);

            var inventory = LayoutReader.Read(reopened);
            Assert.Contains(inventory.Controls, c => c.SdtId == result.ControlId && c.Kind == ControlKind.Label);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertField_afterControl_on_a_real_cell_level_plaintext_control_is_refused_before_any_mutation()
    {
        // YourReference_Lbl is a cell-level SdtCell (parent is a TableRow) AND a bound PLAIN-TEXT control.
        // LayoutEditor used to anchor the new field inside that control's own cell - well-formed OOXML that
        // OpenXmlValidator fully accepts, yet a content control inside a plain-text control, which Word
        // rejects as a corrupt document (only the tool layer's PlainTextNestingGuard caught it). Since GitHub
        // issue #76 the location itself is refused (LocationResolver), so even the raw primitive cannot
        // produce it, and the document is left exactly as it was.
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var body = doc.MainDocumentPart!.Document!.Body!;
            Assert.IsType<SdtCell>(body.Descendants<SdtElement>().Single(s => ReadId(s) == CellLevelLabelId));
            var controlsBefore = body.Descendants<SdtElement>().Count();

            var ex = Assert.Throws<ArgumentException>(() => LayoutEditor.InsertField(
                doc, "/Header/SalesPersonName", new Location { Type = LocationKind.AfterControl, ControlId = CellLevelLabelId }));
            Assert.Equal("location", ex.ParamName);

            Assert.Equal(controlsBefore, body.Descendants<SdtElement>().Count());
            Assert.Empty(PlainTextNestingGuard.Find(doc));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertField_generated_id_does_not_collide_with_any_pre_existing_control_id()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var body = doc.MainDocumentPart!.Document!.Body!;

            // Raw Descendants<SdtElement>() scan - NOT LayoutReader.Read(doc).Controls - so this matches
            // exactly what LayoutEditor.CollectAllControlIds itself avoids colliding with. LayoutReader
            // deliberately does not surface a repeatingSectionItem wrapper as its own inventory entry (it's
            // structural-only - see LayoutReader.Walk), but that wrapper's sdt DOES carry a real w:id that
            // CollectAllControlIds counts and a new id must still avoid; a pre-existing-ids set built from
            // the inventory alone would under-count what actually needs to be avoided and could let this
            // test pass even if uniqueness were only checked against the (smaller) inventory-visible set.
            var preExistingIds = body.Descendants<SdtElement>()
                .Select(ReadId)
                .Where(id => id.HasValue)
                .Select(id => id!.Value)
                .ToHashSet();
            Assert.NotEmpty(preExistingIds); // sanity: the corpus does have pre-existing ids to collide with.

            var result = LayoutEditor.InsertField(doc, FieldPath, new Location { Type = LocationKind.DocumentEnd });

            Assert.DoesNotContain(result.ControlId, preExistingIds);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- LayoutPart (header/footer) targeting ----

    [Fact]
    public void InsertField_targeting_LayoutPart_Header_lands_in_a_header_part_and_passes_validation()
    {
        var path = Corpus.UnlinkedCopy(Corpus.SalesInvoice);
        try
        {
            string expectedPart;
            EditResult result;
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                // header1.xml is this layout's SECTION DEFAULT header. Naming it rather than taking
                // HeaderParts.First() (which is header2.xml here) matters: deriving the expectation from
                // the same collection the resolver used to index into made this test agree with whatever
                // the resolver did, including putting content in the wrong header.
                expectedPart = "header1.xml";

                result = LayoutEditor.InsertField(
                    doc, FieldPath, new Location { Type = LocationKind.DocumentEnd, Part = LayoutPart.Header });

                doc.MainDocumentPart!.Document!.Save();
                foreach (var header in doc.MainDocumentPart.HeaderParts)
                {
                    header.Header?.Save();
                }
            }

            Assert.Equal("insert_field", result.Operation);
            Assert.Equal("Field", result.Kind);
            Assert.Equal(expectedPart, result.Part);
            Assert.Contains($" in {expectedPart}", result.Summary);

            using var reopened = WordprocessingDocument.Open(path, false);
            var openXmlErrors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened).ToList();
            Assert.Empty(openXmlErrors);

            var quick = LayoutValidator.Quick(reopened);
            Assert.Equal(0, quick.ErrorCount);

            var inventory = LayoutReader.Read(reopened);
            Assert.Contains(inventory.Controls, c =>
                c.SdtId == result.ControlId &&
                c.Kind == ControlKind.Field &&
                c.Part == expectedPart);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertLabel_targeting_LayoutPart_Footer_lands_in_a_footer_part_and_passes_validation()
    {
        var path = Corpus.UnlinkedCopy(Corpus.SalesInvoice);
        try
        {
            string expectedPart;
            EditResult result;
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                // footer1.xml is the section default; FooterParts.First() is footer2.xml - see the
                // header test above for why this is named rather than derived.
                expectedPart = "footer1.xml";

                result = LayoutEditor.InsertLabel(
                    doc, LabelPath, new Location { Type = LocationKind.DocumentEnd, Part = LayoutPart.Footer });

                doc.MainDocumentPart!.Document!.Save();
                foreach (var footer in doc.MainDocumentPart.FooterParts)
                {
                    footer.Footer?.Save();
                }
            }

            Assert.Equal("insert_label", result.Operation);
            Assert.Equal("Label", result.Kind);
            Assert.Equal(expectedPart, result.Part);

            using var reopened = WordprocessingDocument.Open(path, false);
            var openXmlErrors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened).ToList();
            Assert.Empty(openXmlErrors);

            var quick = LayoutValidator.Quick(reopened);
            Assert.Equal(0, quick.ErrorCount);

            var inventory = LayoutReader.Read(reopened);
            Assert.Contains(inventory.Controls, c =>
                c.SdtId == result.ControlId &&
                c.Kind == ControlKind.Label &&
                c.Part == expectedPart);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- B28: header/footer scaffolding on demand ----

    [Fact]
    public void InsertLabel_at_documentEnd_in_a_layout_with_no_footer_part_scaffolds_one_and_lands_in_it()
    {
        // A synthetic layout has no header/footer part at all — the shape a from-scratch layout used to be
        // permanently stuck in: layoutPart='footer' could never resolve, so a per-page legal/contact block
        // had to be faked in the body instead.
        var path = SyntheticLayout.Create(SyntheticLayout.PlainParagraph("x"));
        try
        {
            EditResult result;
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                Assert.Empty(doc.MainDocumentPart!.FooterParts);

                result = LayoutEditor.InsertField(
                    doc, "/Header/CompanyName", new Location { Type = LocationKind.DocumentEnd, Part = LayoutPart.Footer });

                doc.MainDocumentPart!.Document!.Save();
                foreach (var footer in doc.MainDocumentPart.FooterParts)
                {
                    footer.Footer?.Save();
                }
            }

            Assert.Contains("footer", result.Part, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("an empty", result.Summary, StringComparison.Ordinal);

            using var reopened = WordprocessingDocument.Open(path, false);
            var main = reopened.MainDocumentPart!;
            var footerPart = Assert.Single(main.FooterParts);

            // The control is really in the new part, and the new part is really wired into the page setup —
            // a footer part nothing references renders on no page.
            Assert.Contains(footerPart.Footer!.Descendants<SdtElement>(), s => ReadId(s) == result.ControlId);
            var sectPr = main.Document!.Body!.Elements<SectionProperties>().Last();
            Assert.Equal(main.GetIdOfPart(footerPart), sectPr.Elements<FooterReference>().Single().Id!.Value);

            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertField_into_a_missing_header_part_scaffolds_only_for_documentEnd_not_for_atText()
    {
        // atText/afterControl/tableCell could not resolve inside a freshly scaffolded (empty) part anyway,
        // so they keep the honest not_found instead of silently gaining a part as a side effect.
        var path = SyntheticLayout.Create(SyntheticLayout.PlainParagraph("x"));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);

            var ex = Assert.Throws<NotFoundException>(() => LayoutEditor.InsertField(
                doc, "/Header/CompanyName",
                new Location { Type = LocationKind.AtText, SearchText = "x", Part = LayoutPart.Header }));

            Assert.Equal(NotFoundTarget.HeaderFooterParts, ex.TargetKind);
            Assert.Empty(doc.MainDocumentPart!.HeaderParts);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertField_targeting_an_unknown_header_PartName_throws_NotFoundException_and_leaves_the_open_document_unmodified()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var sdtCountBefore = doc.MainDocumentPart!.Document!.Descendants<SdtElement>().Count();

            var ex = Assert.Throws<NotFoundException>(() => LayoutEditor.InsertField(
                doc, FieldPath,
                new Location { Type = LocationKind.DocumentEnd, Part = LayoutPart.Header, PartName = "no-such-header.xml" }));

            Assert.Contains("no-such-header.xml", ex.Message);
            Assert.Equal(NotFoundTarget.NamedHeaderFooterPart, ex.TargetKind);
            Assert.Equal(sdtCountBefore, doc.MainDocumentPart!.Document!.Descendants<SdtElement>().Count());
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertField_bad_dataset_path_throws_ArgumentException_and_leaves_the_open_document_unmodified()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var body = doc.MainDocumentPart!.Document!.Body!;
            var sdtCountBefore = body.Descendants<SdtElement>().Count();

            Assert.Throws<ArgumentException>(() =>
                LayoutEditor.InsertField(doc, "/Header/ThisFieldDoesNotExistAnywhere", new Location { Type = LocationKind.DocumentEnd }));

            // SdtFactory validates the dataset path before building anything, so the throw happens before
            // LocationResolver/InsertInline ever run - the in-memory tree must be exactly as it was.
            Assert.Equal(sdtCountBefore, body.Descendants<SdtElement>().Count());
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- remove_control ----

    [Fact]
    public void RemoveControl_without_keepText_removes_the_control_and_its_text_entirely()
    {
        // Synthetic (not corpus): guarantees the placeholder text is unique in the document, so "the text
        // is gone" is a meaningful assertion — a corpus layout can legitimately contain more than one
        // control sharing the same unmerged placeholder text (e.g. two controls both bound to the same
        // field), which would make that assertion unreliable there.
        var path = SyntheticLayout.Create(SyntheticLayout.InlineControlWithId(200, "REMOVE-ENTIRELY-UNIQUE-TEXT"));
        try
        {
            EditResult removeResult;
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var body = doc.MainDocumentPart!.Document!.Body!;
                Assert.Contains(body.Descendants<SdtElement>(), s => ReadId(s) == 200);

                removeResult = LayoutEditor.RemoveControl(doc, 200, keepText: false);
                doc.MainDocumentPart!.Document!.Save();
            }

            Assert.Equal("remove_control", removeResult.Operation);
            Assert.Equal(200, removeResult.ControlId);
            Assert.Equal("document.xml", removeResult.Part);

            using var reopened = WordprocessingDocument.Open(path, false);
            var openXmlErrors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened).ToList();
            Assert.Empty(openXmlErrors);

            var reopenedBody = reopened.MainDocumentPart!.Document!.Body!;
            Assert.DoesNotContain(reopenedBody.Descendants<SdtElement>(), s => ReadId(s) == 200);
            Assert.DoesNotContain(reopenedBody.Descendants<Text>(), t => t.Text == "REMOVE-ENTIRELY-UNIQUE-TEXT");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RemoveControl_unknown_id_throws_NotFoundException_with_Control_target()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);

            var ex = Assert.Throws<NotFoundException>(() => LayoutEditor.RemoveControl(doc, 999_999_999, keepText: false));
            Assert.Contains("999999999", ex.Message);
            Assert.Equal(NotFoundTarget.Control, ex.TargetKind);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RemoveControl_finds_and_removes_a_control_living_in_a_header_or_footer_part()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            int headerControlId;
            string expectedPart;
            using (var probe = WordprocessingDocument.Open(path, false))
            {
                var headerOrFooterControl = LayoutReader.Read(probe).Controls
                    .First(c => c.SdtId.HasValue && !string.Equals(c.Part, "document.xml", StringComparison.Ordinal));
                headerControlId = headerOrFooterControl.SdtId!.Value;
                expectedPart = headerOrFooterControl.Part;
            }

            EditResult removeResult;
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                removeResult = LayoutEditor.RemoveControl(doc, headerControlId, keepText: false);
                doc.MainDocumentPart!.Document!.Save();
                foreach (var header in doc.MainDocumentPart.HeaderParts)
                {
                    header.Header?.Save();
                }

                foreach (var footer in doc.MainDocumentPart.FooterParts)
                {
                    footer.Footer?.Save();
                }
            }

            Assert.Equal(expectedPart, removeResult.Part);

            using var reopened = WordprocessingDocument.Open(path, false);
            var stillPresent = LayoutReader.Read(reopened).Controls.Any(c => c.SdtId == headerControlId);
            Assert.False(stillPresent, $"control {headerControlId} should have been removed from '{expectedPart}'");

            var openXmlErrors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened).ToList();
            Assert.Empty(openXmlErrors);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- remove_control keepText: generic across the four concrete sdt kinds (synthetic fixtures) ----

    [Fact]
    public void RemoveControl_keepText_on_an_inline_control_keeps_its_run_in_the_same_paragraph()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.InlineControlWithId(55, "KEEP-INLINE"));
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var body = doc.MainDocumentPart!.Document!.Body!;
                var paragraph = body.Descendants<SdtElement>().Single(s => ReadId(s) == 55).Ancestors<Paragraph>().First();

                LayoutEditor.RemoveControl(doc, 55, keepText: true);

                Assert.DoesNotContain(body.Descendants<SdtElement>(), s => ReadId(s) == 55);
                Assert.Contains(paragraph.Elements<Run>(), r => r.GetFirstChild<Text>()?.Text == "KEEP-INLINE");
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var openXmlErrors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened).ToList();
            Assert.Empty(openXmlErrors);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RemoveControl_keepText_on_a_block_control_keeps_its_paragraph_as_a_direct_body_child()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.BlockControlWithId(77, "KEEP-BLOCK"));
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var body = doc.MainDocumentPart!.Document!.Body!;

                LayoutEditor.RemoveControl(doc, 77, keepText: true);

                Assert.DoesNotContain(body.Descendants<SdtElement>(), s => ReadId(s) == 77);
                Assert.Contains(body.Elements<Paragraph>(), p => p.Descendants<Text>().Any(t => t.Text == "KEEP-BLOCK"));
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var openXmlErrors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened).ToList();
            Assert.Empty(openXmlErrors);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- remove_control keepText on the w:group of a block-level top-level repeater (GitHub issue #87) ----

    private const string TopLevelHeaderXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header";
    private const string NestedLineXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line";

    [Fact]
    public void RemoveControl_keepText_on_the_group_of_a_top_level_block_repeater_throws_and_leaves_it_in_place()
    {
        // Unwrapping the group turns StandardStatement's shape (repeatingSection > group > item) into the one
        // BC prints as an EMPTY document - exactly what the top-level-repeater-group check reports.
        var path = SyntheticLayout.Create(SyntheticLayout.GroupedRepeater(TopLevelHeaderXPath, SyntheticLayout.GoodItemId));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);

            var ex = Assert.Throws<ArgumentException>(() => LayoutEditor.RemoveControl(doc, 111, keepText: true));

            Assert.Equal("keepText", ex.ParamName);
            Assert.Contains("w:group", ex.Message, StringComparison.Ordinal);
            Assert.Contains("id 110", ex.Message, StringComparison.Ordinal);
            Assert.Contains(doc.MainDocumentPart!.Document!.Body!.Descendants<SdtElement>(), s => ReadId(s) == 111);
            Assert.DoesNotContain(LayoutValidator.Quick(doc).Findings, f => f.Check == "top-level-repeater-group");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RemoveControl_keepText_on_the_group_of_a_nested_item_block_repeater_still_unwraps_it()
    {
        // Only the top-level shape needs the group (block repeaters over nested items render without one in
        // the stock 1304/1306 layouts), so the refusal must not reach any further than the check does.
        var path = SyntheticLayout.Create(
            SyntheticLayout.GroupedRepeater(NestedLineXPath, SyntheticLayout.GoodItemId),
            datasetXml: SyntheticLayout.DatasetWithLineXml);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);

            LayoutEditor.RemoveControl(doc, 111, keepText: true);

            var body = doc.MainDocumentPart!.Document!.Body!;
            Assert.DoesNotContain(body.Descendants<SdtElement>(), s => ReadId(s) == 111);
            Assert.Contains(body.Descendants<SdtElement>(), s => ReadId(s) == 112);
            Assert.DoesNotContain(LayoutValidator.Quick(doc).Findings, f => f.Severity == FindingSeverity.Error);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RemoveControl_keepText_on_another_group_inside_a_top_level_repeaters_item_still_unwraps_it()
    {
        // The refusal asks whether the unwrap would take away the repeater's group, not whether the control
        // is "a group in a top-level repeater": a group INSIDE the item is not the one BC needs.
        var body =
            "<w:sdt><w:sdtPr><w:id w:val=\"110\"/>"
            + $"<w15:dataBinding w:xpath=\"{TopLevelHeaderXPath}\" w:storeItemID=\"{SyntheticLayout.GoodItemId}\"/>"
            + "<w15:repeatingSection/></w:sdtPr><w:sdtContent>"
            + "<w:sdt><w:sdtPr><w:id w:val=\"111\"/><w:lock w:val=\"contentLocked\"/><w:group/></w:sdtPr><w:sdtContent>"
            + "<w:sdt><w:sdtPr><w:id w:val=\"112\"/><w15:repeatingSectionItem/></w:sdtPr><w:sdtContent>"
            + "<w:sdt><w:sdtPr><w:id w:val=\"113\"/><w:group/></w:sdtPr><w:sdtContent>"
            + "<w:p><w:r><w:t>inner</w:t></w:r></w:p></w:sdtContent></w:sdt>"
            + "</w:sdtContent></w:sdt>"
            + "</w:sdtContent></w:sdt>"
            + "</w:sdtContent></w:sdt>";
        var path = SyntheticLayout.Create(body);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);

            LayoutEditor.RemoveControl(doc, 113, keepText: true);

            Assert.DoesNotContain(doc.MainDocumentPart!.Document!.Body!.Descendants<SdtElement>(), s => ReadId(s) == 113);
            Assert.DoesNotContain(LayoutValidator.Quick(doc).Findings, f => f.Check == "top-level-repeater-group");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RemoveControl_keepText_on_a_cell_level_control_keeps_a_plain_TableCell_in_the_row()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.TableWithCellLevelControl(id: 99, text: "KEEP-CELL"));
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var body = doc.MainDocumentPart!.Document!.Body!;
                var row = body.Descendants<TableRow>().Single();

                LayoutEditor.RemoveControl(doc, 99, keepText: true);

                Assert.DoesNotContain(body.Descendants<SdtElement>(), s => ReadId(s) == 99);
                Assert.Empty(row.Elements<SdtCell>());
                Assert.Equal(2, row.Elements<TableCell>().Count());
                Assert.Contains(row.Elements<TableCell>(), c => c.Descendants<Text>().Any(t => t.Text == "KEEP-CELL"));
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var openXmlErrors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened).ToList();
            Assert.Empty(openXmlErrors);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RemoveControl_keepText_false_on_a_cell_level_control_keeps_the_column_but_empties_it()
    {
        // The bug this guards against: a cell-level control (SdtCell) wraps a whole w:tc = one column.
        // keepText=false must NOT delete that cell (which would leave the row with fewer cells than the
        // grid and break the table); it preserves the cell and only empties its text.
        var path = SyntheticLayout.Create(SyntheticLayout.TableWithCellLevelControl(id: 99, text: "DROP-CELL-TEXT"));
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var body = doc.MainDocumentPart!.Document!.Body!;
                var row = body.Descendants<TableRow>().Single();

                var result = LayoutEditor.RemoveControl(doc, 99, keepText: false);

                // The sdt wrapper is gone, but the column is preserved: still two cells in the row.
                Assert.DoesNotContain(body.Descendants<SdtElement>(), s => ReadId(s) == 99);
                Assert.Empty(row.Elements<SdtCell>());
                Assert.Equal(2, row.Elements<TableCell>().Count());
                // The former field text is gone (keepText=false), but the (now empty) cell still has a
                // paragraph so it stays well-formed.
                Assert.DoesNotContain(row.Descendants<Text>(), t => t.Text == "DROP-CELL-TEXT");
                Assert.All(row.Elements<TableCell>(), c => Assert.NotEmpty(c.Elements<Paragraph>()));
                Assert.Contains("column is preserved", result.Summary);
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var openXmlErrors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened).ToList();
            Assert.Empty(openXmlErrors);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RemoveControl_keepText_false_on_a_block_control_that_is_a_cells_sole_content_keeps_a_valid_cell()
    {
        // Regression for the reported corruption: a line-items amount/quantity field is a BLOCK-level sdt
        // that is the SOLE content of its w:tc. Removing it with keepText=false used to take the cell's only
        // paragraph with it, leaving an empty <w:tc> (just its w:tcPr) — a silently-corrupt document (Word
        // rejects it; OpenXmlValidator does not). The cell must be left with a paragraph so it stays valid.
        var path = SyntheticLayout.Create(SyntheticLayout.TableWithBlockControlInCell(id: 88, text: "DROP-BLOCK"));
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var body = doc.MainDocumentPart!.Document!.Body!;
                var row = body.Descendants<TableRow>().Single();

                var result = LayoutEditor.RemoveControl(doc, 88, keepText: false);

                Assert.DoesNotContain(body.Descendants<SdtElement>(), s => ReadId(s) == 88);
                // The column is preserved: the row still has two cells, and the (now empty) cell that held
                // the control still has a paragraph, so it stays well-formed rather than becoming an empty tc.
                Assert.Equal(2, row.Elements<TableCell>().Count());
                Assert.DoesNotContain(row.Descendants<Text>(), t => t.Text == "DROP-BLOCK");
                Assert.All(row.Elements<TableCell>(), c => Assert.NotEmpty(c.Elements<Paragraph>()));
                Assert.Contains("column is preserved", result.Summary);
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var openXmlErrors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened).ToList();
            Assert.Empty(openXmlErrors);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RemoveControl_keepText_false_on_a_corpus_address_cell_preserves_the_rows_column_count()
    {
        // Regression for the reported bug: removing CustomerAddress6 (a cell-level control in a 3-column
        // header row) used to delete the whole column and break the table. It must now leave the row with
        // the SAME number of cell-bearing children (2 sibling SdtCells + 1 plain TableCell where the removed
        // control's cell was = still 3), so the grid stays intact.
        const int customerAddress6Id = -2064325541; // #Nav: /Header/CustomerAddress6 in SalesInvoiceForSubscriptionBilling.docx
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            int cellsBefore;
            using (var probe = WordprocessingDocument.Open(path, false))
            {
                var sdt = probe.MainDocumentPart!.Document!.Descendants<SdtElement>().Single(s => ReadId(s) == customerAddress6Id);
                Assert.IsType<SdtCell>(sdt);
                var row = sdt.Ancestors<TableRow>().First();
                cellsBefore = row.ChildElements.Count(e => e is TableCell or SdtCell);
            }

            using (var doc = WordprocessingDocument.Open(path, true))
            {
                LayoutEditor.RemoveControl(doc, customerAddress6Id, keepText: false);
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            Assert.DoesNotContain(reopened.MainDocumentPart!.Document!.Descendants<SdtElement>(), s => ReadId(s) == customerAddress6Id);

            // The former control's cell is now a plain TableCell sitting among its untouched sibling cells:
            // the row still has exactly as many cell-bearing children as before.
            var companyAddress6 = reopened.MainDocumentPart!.Document!.Descendants<SdtElement>()
                .First(s => ReadAlias(s) == "#Nav: /Header/CompanyAddress6");
            var rowAfter = companyAddress6.Ancestors<TableRow>().First();
            Assert.Equal(cellsBefore, rowAfter.ChildElements.Count(e => e is TableCell or SdtCell));

            var openXmlErrors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened).ToList();
            Assert.Empty(openXmlErrors);
        }
        finally
        {
            File.Delete(path);
        }
    }

    private static string? ReadAlias(SdtElement sdt) =>
        sdt.GetFirstChild<SdtProperties>()?.GetFirstChild<SdtAlias>()?.Val?.Value;

    [Fact]
    public void RemoveControl_keepText_on_a_row_level_control_keeps_a_plain_TableRow_in_the_table()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.TableWithRowLevelControl(id: 111, text: "KEEP-ROW"));
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var body = doc.MainDocumentPart!.Document!.Body!;
                var table = body.Descendants<Table>().Single();

                LayoutEditor.RemoveControl(doc, 111, keepText: true);

                Assert.DoesNotContain(body.Descendants<SdtElement>(), s => ReadId(s) == 111);
                Assert.Empty(table.Elements<SdtRow>());
                Assert.Contains(table.Elements<TableRow>(), r => r.Descendants<Text>().Any(t => t.Text == "KEEP-ROW"));
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var openXmlErrors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened).ToList();
            Assert.Empty(openXmlErrors);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RemoveControl_keepText_false_on_a_tables_only_row_level_control_removes_the_table_too()
    {
        // whole-removal of a row-level control (SdtRow) that is a table's ONLY row must not leave a
        // rowless <w:tbl> behind — schema-legal but Word-hostile dead markup. The table itself must go.
        var path = SyntheticLayout.Create(SyntheticLayout.TableWithRowLevelControl(id: 111, text: "ONLY-ROW"));
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var body = doc.MainDocumentPart!.Document!.Body!;

                var result = LayoutEditor.RemoveControl(doc, 111, keepText: false);

                Assert.DoesNotContain(body.Descendants<SdtElement>(), s => ReadId(s) == 111);
                Assert.Empty(body.Descendants<Table>());
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var openXmlErrors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened).ToList();
            Assert.Empty(openXmlErrors);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RemoveControl_keepText_false_removing_a_nested_tables_only_row_preserves_the_outer_cell()
    {
        // Removing a nested table's own last row-level control removes
        // that INNER table (as the sibling test above proves) - but the inner table can itself be the SOLE
        // content of an OUTER table's cell (a nested data table - a real BC shape). Removing it wholesale
        // would then repeat the exact same hazard one level up: an empty outer <w:tc> (just its w:tcPr, no
        // paragraph) - schema-legal to OpenXmlValidator but Word-corrupting. The outer cell must survive as
        // a valid empty cell, exactly like the plain cell-level-control case.
        const int InnerRowControlId = 222;
        var innerTable =
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid>"
            + "<w:sdt><w:sdtPr>"
            + $"<w:id w:val=\"{InnerRowControlId}\"/>"
            + "</w:sdtPr><w:sdtContent><w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>NESTED-ROW</w:t></w:r></w:p></w:tc></w:tr></w:sdtContent></w:sdt>"
            + "</w:tbl>";
        // Deliberately NO trailing w:p after the nested w:tbl in the outer cell - the exact Word-hostile-
        // but-validator-silent shape this guard exists for (same pattern as SyntheticLayout's own
        // TableWithBlockControlInCell fixture).
        var outerBody =
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"4000\"/></w:tblGrid>"
            + $"<w:tr><w:tc><w:tcPr/>{innerTable}</w:tc></w:tr>"
            + "</w:tbl>";
        var path = SyntheticLayout.Create(outerBody);
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var body = doc.MainDocumentPart!.Document!.Body!;

                var result = LayoutEditor.RemoveControl(doc, InnerRowControlId, keepText: false);

                // The inner (nested) table is gone entirely...
                Assert.DoesNotContain(body.Descendants<SdtElement>(), s => ReadId(s) == InnerRowControlId);
                var remainingTable = Assert.Single(body.Descendants<Table>());

                // ...and the OUTER cell that held it survives as a valid, non-empty cell (a w:p was added
                // back), not a bare w:tcPr with no block content.
                var outerCell = Assert.Single(remainingTable.Descendants<TableCell>());
                Assert.Contains(outerCell.Elements<Paragraph>(), _ => true);

                doc.MainDocumentPart!.Document!.Save();

                // Both custom guards agree the saved document is structurally sound.
                Assert.Empty(TableGridConsistencyGuard.Find(doc));
                var quick = LayoutValidator.Quick(doc);
                Assert.True(quick.Passed, string.Join(" | ", quick.Findings.Select(f => f.Message)));
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var openXmlErrors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened).ToList();
            Assert.Empty(openXmlErrors);
            Assert.Empty(TableGridConsistencyGuard.Find(reopened));

            var reopenedQuick = LayoutValidator.Quick(path);
            Assert.True(reopenedQuick.Passed, string.Join(" | ", reopenedQuick.Findings.Select(f => f.Message)));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RemoveControl_keepText_false_on_one_row_of_a_multi_row_table_leaves_the_table_in_place()
    {
        // Symmetric guard: only a NOW-EMPTY table should be removed. A table with other rows remaining
        // must keep those rows and stay a <w:tbl>.
        var fragment = SyntheticLayout.TableWithRowLevelControl(id: 111, text: "ROW-A")
            .Replace("</w:tbl>", "<w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>ROW-B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>");
        var path = SyntheticLayout.Create(fragment);
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var body = doc.MainDocumentPart!.Document!.Body!;

                LayoutEditor.RemoveControl(doc, 111, keepText: false);

                var table = Assert.Single(body.Descendants<Table>());
                Assert.DoesNotContain(body.Descendants<SdtElement>(), s => ReadId(s) == 111);
                Assert.Contains(table.Descendants<Text>(), t => t.Text == "ROW-B");
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var openXmlErrors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened).ToList();
            Assert.Empty(openXmlErrors);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RemoveControl_keepText_preserves_non_run_children_alongside_the_run()
    {
        // Real BC layouts sometimes wrap a control's run in w:proofErr spell-check markers (confirmed in
        // the corpus itself - see InlineControlWithProofErr's remarks); UnwrapSdt must move ALL of the
        // content's children, not just the one that looks like "the real content", or these markers - and
        // anything else riding alongside the run - would be silently dropped.
        var path = SyntheticLayout.Create(SyntheticLayout.InlineControlWithProofErr(321, "KEEP-WITH-PROOFERR"));
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var body = doc.MainDocumentPart!.Document!.Body!;
                var paragraph = body.Descendants<SdtElement>().Single(s => ReadId(s) == 321).Ancestors<Paragraph>().First();

                LayoutEditor.RemoveControl(doc, 321, keepText: true);

                Assert.DoesNotContain(body.Descendants<SdtElement>(), s => ReadId(s) == 321);
                // Both proofErr markers AND the run must have survived, as direct children of the paragraph
                // (not just the run alone).
                Assert.Equal(2, paragraph.Elements<ProofError>().Count());
                Assert.Contains(paragraph.Elements<Run>(), r => r.GetFirstChild<Text>()?.Text == "KEEP-WITH-PROOFERR");
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var openXmlErrors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened).ToList();
            Assert.Empty(openXmlErrors);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- remove_control keepText rejection: a repeating-section control must not be unwrapped ----

    [Fact]
    public void RemoveControl_keepText_true_on_a_row_repeater_throws_ArgumentException_and_changes_nothing()
    {
        // A ROW-level repeater: keepText=true unwraps a block-level one instead (wrap_in_repeater's undo,
        // GitHub issue #31 - see WrapInRepeaterTests).
        var path = SyntheticLayout.Create(SyntheticLayout.RepeaterTable(
            "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line",
            "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line[1]/ns0:No[1]", SyntheticLayout.GoodItemId));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var body = doc.MainDocumentPart!.Document!.Body!;

            var ex = Assert.Throws<ArgumentException>(() => LayoutEditor.RemoveControl(doc, 700, keepText: true));
            Assert.Contains("repeating-section", ex.Message);

            // Rejected before any mutation: the repeater (id 700) and its item (id 701) are both still
            // present, still nested exactly as they were.
            Assert.Contains(body.Descendants<SdtElement>(), s => ReadId(s) == 700);
            Assert.Contains(body.Descendants<SdtElement>(), s => ReadId(s) == 701);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RemoveControl_keepText_false_on_a_repeater_control_still_removes_the_whole_repeater()
    {
        // Only the keepText=true unwrap is rejected - whole-control removal of a repeater stays allowed.
        var path = SyntheticLayout.Create(SyntheticLayout.ProperRepeater(
            "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line", SyntheticLayout.GoodItemId));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var body = doc.MainDocumentPart!.Document!.Body!;

            var result = LayoutEditor.RemoveControl(doc, 100, keepText: false);

            Assert.Equal("Repeater", result.Kind);
            Assert.DoesNotContain(body.Descendants<SdtElement>(), s => ReadId(s) == 100);
            Assert.DoesNotContain(body.Descendants<SdtElement>(), s => ReadId(s) == 101);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- insert_text: static inline runs ----

    [Fact]
    public void InsertText_appends_a_plain_run_with_no_content_control()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            EditResult result;
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var controlsBefore = LayoutReader.Read(doc).Controls.Count;
                result = LayoutEditor.InsertText(doc, "Page ", new Location { Type = LocationKind.DocumentEnd });
                doc.MainDocumentPart!.Document!.Save();

                // The whole point: no new content control appears in the inventory.
                Assert.Equal(controlsBefore, LayoutReader.Read(doc).Controls.Count);
            }

            Assert.Equal("insert_text", result.Operation);
            Assert.Equal(0, result.ControlId);
            Assert.Equal("StaticText", result.Kind);
            Assert.Equal("document.xml", result.Part);

            using var reopened = WordprocessingDocument.Open(path, false);
            var body = reopened.MainDocumentPart!.Document!.Body!;
            var texts = body.Descendants<Text>().Select(t => t.Text).ToList();
            Assert.Contains("Page ", texts);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertText_preserves_a_whitespace_only_run()
    {
        // The separator case this tool exists for. Without xml:space="preserve" Word drops the space and
        // silently undoes the call, which is exactly the "Document NoDOCU-0150" symptom being fixed.
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                LayoutEditor.InsertText(doc, " ", new Location { Type = LocationKind.DocumentEnd });
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var spaceRun = reopened.MainDocumentPart!.Document!.Body!
                .Descendants<Text>()
                .FirstOrDefault(t => t.Text == " ");
            Assert.NotNull(spaceRun);
            Assert.Equal(SpaceProcessingModeValues.Preserve, spaceRun!.Space!.Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertText_afterControl_lands_in_the_same_paragraph_as_the_control()
    {
        // The reason the tool exists: gluing two inline controls together needs the run to sit BETWEEN them
        // in one paragraph, not in a paragraph of its own.
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            int labelId;
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                // Anchor on a freshly inserted INLINE control (an SdtRun). A cell-level control would not
                // prove the point: LocationResolver deliberately anchors inside such a control's own cell
                // rather than beside it, because a w:tr cannot host a run as a direct child.
                labelId = LayoutEditor
                    .InsertLabel(doc, LabelPath, new Location { Type = LocationKind.DocumentEnd })
                    .ControlId;
                LayoutEditor.InsertText(doc, ": ",
                    new Location { Type = LocationKind.AfterControl, ControlId = labelId });
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var control = reopened.MainDocumentPart!.Document!.Body!
                .Descendants<SdtElement>()
                .First(s => ReadId(s) == labelId);
            var sibling = control.NextSibling();
            Assert.IsType<Run>(sibling);
            Assert.Equal(": ", ((Run)sibling!).GetFirstChild<Text>()!.Text);

            // Same paragraph as the control - the whole reason the tool exists (a run in a paragraph of its
            // own would put the separator on another line).
            Assert.IsType<Paragraph>(control.Parent);
            Assert.Same(control.Parent, sibling!.Parent);

            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertText_applies_bold_and_font_size_to_the_run()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                LayoutEditor.InsertText(doc, "Total", new Location { Type = LocationKind.DocumentEnd },
                    new TextFormat { Bold = true, FontSizePoints = 8 });
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var run = reopened.MainDocumentPart!.Document!.Body!
                .Descendants<Run>()
                .First(r => r.GetFirstChild<Text>()?.Text == "Total");
            Assert.NotNull(run.RunProperties?.Bold);
            Assert.Equal("16", run.RunProperties!.FontSize!.Val!.Value); // 8 pt = 16 half-points
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertText_rejects_empty_text_but_not_a_space()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var ex = Assert.Throws<ArgumentException>(
                () => LayoutEditor.InsertText(doc, "", new Location { Type = LocationKind.DocumentEnd }));
            Assert.Contains("non-empty", ex.Message);

            // A space must NOT be swept up by the same guard - it is the primary use case.
            LayoutEditor.InsertText(doc, " ", new Location { Type = LocationKind.DocumentEnd });
        }
        finally
        {
            File.Delete(path);
        }
    }

}
