using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;
using BcWordLayout.McpHost;
using BcWordLayout.McpHost.Tools;
using BcWordLayout.Merge;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Blip = DocumentFormat.OpenXml.Drawing.Blip;
using DocProperties = DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties;
using Extent = DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent;

namespace BcWordLayout.Tests;

/// <summary>
/// Covers <see cref="SdtFactory.BuildPicture"/> / <see cref="LayoutEditor.InsertPicture"/> and the
/// <c>insert_picture</c> tool. The shape is mirrored from the real add-in-authored
/// <c>/Header/CompanyPicture</c> control in <c>tests/corpus/StandardSalesQuote.docx</c>'s
/// <c>header3.xml</c>; these tests author against <c>SalesInvoiceForSubscriptionBilling.docx</c>, whose
/// dataset declares the same <c>CompanyPicture</c> column. The two proofs that matter: the layout is still
/// valid to <see cref="OpenXmlValidator"/> (a dangling image reference or a duplicate <c>wp:docPr</c> id
/// would not be), and <see cref="MergeEngine"/> recognizes and FILLS the authored placeholder exactly like
/// a BC-authored one.
/// </summary>
public class PictureTests
{
    private const string PicturePath = "/Header/CompanyPicture";

    private static string CopyOfCorpus(string corpusFile)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-picture-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    private static List<ValidationErrorInfo> OpenXmlErrors(WordprocessingDocument doc) =>
        new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc).ToList();

    [Fact]
    public void Tool_insert_picture_adds_a_bound_picture_control_that_reads_back_as_a_Picture()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = EditTools.InsertPicture(path, PicturePath, "documentEnd");
            Assert.True(response.Ok, response.Error?.Message);
            var dto = Assert.IsType<EditResultDto>(response.Data);
            Assert.Equal("insert_picture", dto.Operation);
            Assert.Equal("Picture", dto.Kind);
            Assert.Equal("/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:CompanyPicture[1]", dto.XPath);
            Assert.True(dto.QuickValidation.Passed, "an authored picture must not introduce findings");

            using var reopened = WordprocessingDocument.Open(path, false);
            var inventory = LayoutReader.Read(reopened);
            var control = Assert.Single(inventory.Controls, c => c.SdtId == dto.ControlId);
            Assert.Equal(ControlKind.Picture, control.Kind);
            Assert.Equal("#Nav: /Header/CompanyPicture", control.Alias);

            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertPicture_embeds_a_real_image_part_in_the_hosting_part_that_the_blip_resolves_to()
    {
        // A picture frame whose r:embed points at nothing is a corrupt document to Word (and the
        // relationship is part-scoped, so the image part must live in the SAME part as the control).
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = EditTools.InsertPicture(path, PicturePath, "documentEnd", layoutPart: "header");
            Assert.True(response.Ok, response.Error?.Message);
            var dto = Assert.IsType<EditResultDto>(response.Data);

            using var reopened = WordprocessingDocument.Open(path, false);
            var headerPart = reopened.MainDocumentPart!.HeaderParts
                .Single(h => Path.GetFileName(h.Uri.OriginalString) == dto.Part);

            var sdt = headerPart.Header!.Descendants<SdtElement>()
                .Single(s => s.GetFirstChild<SdtProperties>()?.GetFirstChild<SdtId>()?.Val?.Value == dto.ControlId);

            var embedId = sdt.Descendants<Blip>().Single().Embed!.Value!;
            var imagePart = Assert.IsAssignableFrom<ImagePart>(headerPart.GetPartById(embedId));

            using var stream = imagePart.GetStream();
            var bytes = new byte[8];
            Assert.Equal(8, stream.Read(bytes, 0, 8));
            Assert.Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }, bytes); // PNG signature

            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertPicture_sizes_the_frame_from_widthMm_heightMm_and_defaults_to_the_corpus_30mm_square()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var defaulted = EditTools.InsertPicture(path, PicturePath, "documentEnd");
            var sized = EditTools.InsertPicture(path, PicturePath, "documentEnd", widthMm: 50, heightMm: 20);
            Assert.True(sized.Ok, sized.Error?.Message);

            using var reopened = WordprocessingDocument.Open(path, false);
            var body = reopened.MainDocumentPart!.Document!.Body!;

            var defaultExtent = ExtentOf(body, Assert.IsType<EditResultDto>(defaulted.Data).ControlId);
            Assert.Equal(SdtFactory.CorpusPictureExtentEmu, defaultExtent.Cx!.Value);
            Assert.Equal(SdtFactory.CorpusPictureExtentEmu, defaultExtent.Cy!.Value);

            var sizedExtent = ExtentOf(body, Assert.IsType<EditResultDto>(sized.Data).ControlId);
            Assert.Equal(50 * SdtFactory.EmuPerMillimetre, sizedExtent.Cx!.Value);
            Assert.Equal(20 * SdtFactory.EmuPerMillimetre, sizedExtent.Cy!.Value);

            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Two_inserted_pictures_get_distinct_drawing_ids()
    {
        // wp:docPr/@id must be unique document-wide; a duplicate trips OpenXmlValidator's
        // Sem_UniqueAttributeValue and Word's own repair prompt. The corpus layout already contains a
        // picture of its own (PaymentServiceLogo), so the scan must consider existing drawings too.
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            Assert.True(EditTools.InsertPicture(path, PicturePath, "documentEnd").Ok);
            Assert.True(EditTools.InsertPicture(path, PicturePath, "documentEnd").Ok);

            using var reopened = WordprocessingDocument.Open(path, false);
            var ids = reopened.MainDocumentPart!.Document!.Body!
                .Descendants<DocProperties>()
                .Select(p => p.Id!.Value)
                .ToList();

            Assert.Equal(ids.Count, ids.Distinct().Count());
            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void An_authored_picture_is_filled_by_the_merge_engine_exactly_like_a_BC_authored_one()
    {
        // The end-to-end proof that the authored shape is the real shape: MergeEngine classifies it as a
        // picture control and repoints its blip (PicturesFilled), with no picture-no-blip warning.
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        var mergedPath = Path.Combine(Path.GetTempPath(), $"bcwl-picture-merged-{Guid.NewGuid():N}.docx");
        try
        {
            var before = MergeEngine.Merge(path, mergedPath, new MergeOptions { Rows = 2 });
            Assert.True(EditTools.InsertPicture(path, PicturePath, "documentEnd").Ok);
            var after = MergeEngine.Merge(path, mergedPath, new MergeOptions { Rows = 2 });

            Assert.Equal(before.Stats.PicturesFilled + 1, after.Stats.PicturesFilled);
            Assert.DoesNotContain(after.Warnings, w => w.Kind == "picture-no-blip");

            using var merged = WordprocessingDocument.Open(mergedPath, false);
            Assert.Empty(OpenXmlErrors(merged));
        }
        finally
        {
            File.Delete(path);
            if (File.Exists(mergedPath))
            {
                File.Delete(mergedPath);
            }
        }
    }

    [Fact]
    public void Tool_insert_picture_rejects_an_out_of_range_size_and_a_non_picture_dataset_path()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var tooBig = EditTools.InsertPicture(path, PicturePath, "documentEnd", widthMm: 5000);
            Assert.False(tooBig.Ok);
            Assert.Equal("invalid_argument", tooBig.Error!.Code);
            Assert.Contains("millimetres", tooBig.Error.Hint, StringComparison.OrdinalIgnoreCase);

            var unknownPath = EditTools.InsertPicture(path, "/Header/NoSuchPicture", "documentEnd");
            Assert.False(unknownPath.Ok);
            Assert.Equal("invalid_argument", unknownPath.Error!.Code);
        }
        finally
        {
            File.Delete(path);
        }
    }

    private static Extent ExtentOf(OpenXmlElement root, int controlId) =>
        root.Descendants<SdtElement>()
            .Single(s => s.GetFirstChild<SdtProperties>()?.GetFirstChild<SdtId>()?.Val?.Value == controlId)
            .Descendants<Extent>()
            .Single();
}
