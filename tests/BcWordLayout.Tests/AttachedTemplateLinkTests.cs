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
/// Business Central refuses to upload a layout whose settings link an external attached template
/// ("UnsupportedExternalContentFields: ... ExternalRelationship:attachedTemplate", a BC 28 sandbox, 2026-10-24), and
/// four of Microsoft's own stock layouts in the corpus link one. Covers <see cref="AttachedTemplateLink"/> itself,
/// every tool that writes a layout dropping the link and saying so, and <c>create_layout</c> dropping a template's.
/// The validator's error is pinned in <see cref="LayoutValidatorTests"/>.
/// </summary>
public class AttachedTemplateLinkTests
{
    private const string RelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string AttachedTemplateRelationshipType = RelationshipsNamespace + "/attachedTemplate";
    private const string ImageRelationshipType = RelationshipsNamespace + "/image";

    /// <summary>The target SalesInvoiceForSubscriptionBilling's settings link, as the corpus scrub left it.</summary>
    private const string CorpusTemplate = "file:///C:/Templates/Simplified Sales Document.dotx";

    private static string CopyOfCorpus(string corpusFile)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-attached-template-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    private static string TempPath() => Path.Combine(Path.GetTempPath(), $"bcwl-attached-template-{Guid.NewGuid():N}.docx");

    /// <summary>Every external attached-template relationship and <c>w:attachedTemplate</c> element in the file.</summary>
    private static (int Relationships, int Settings) Links(string path)
    {
        using var doc = WordprocessingDocument.Open(path, false);
        var relationships = AttachedTemplateLink.Find(doc.MainDocumentPart!).Count();
        var settings = doc.MainDocumentPart!.DocumentSettingsPart?.Settings?.Elements<AttachedTemplate>().Count() ?? 0;
        return (relationships, settings);
    }

    /// <summary>Adds an external attached-template link to the settings part of the layout at <paramref name="path"/>.</summary>
    private static void LinkTemplate(string path, string target)
    {
        using var doc = WordprocessingDocument.Open(path, true);
        var main = doc.MainDocumentPart!;
        var settingsPart = main.DocumentSettingsPart ?? main.AddNewPart<DocumentSettingsPart>();
        settingsPart.Settings ??= new Settings();
        var relationship = settingsPart.AddExternalRelationship(AttachedTemplateRelationshipType, new Uri(target));
        settingsPart.Settings.PrependChild(new AttachedTemplate { Id = relationship.Id });
        settingsPart.Settings.Save();
    }

    [Fact]
    public void Remove_unlinks_the_template_and_its_setting_and_leaves_every_other_external_link()
    {
        var path = TempPath();
        try
        {
            using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                main.Document = new Document(new Body(new Paragraph(new Run(new Text("x")))));
                main.AddExternalRelationship(ImageRelationshipType, new Uri("https://example.com/logo.png"));
                main.AddHyperlinkRelationship(new Uri("https://example.com/"), true);
            }

            LinkTemplate(path, @"\\templates-host\share\letterhead.dotx");

            IReadOnlyList<string> unlinked;
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                unlinked = AttachedTemplateLink.Remove(doc);
            }

            Assert.Equal("file://templates-host/share/letterhead.dotx", Assert.Single(unlinked));
            Assert.Equal((0, 0), Links(path));
            using (var doc = WordprocessingDocument.Open(path, false))
            {
                var main = doc.MainDocumentPart!;
                Assert.Single(main.ExternalRelationships, r => r.RelationshipType == ImageRelationshipType);
                Assert.Single(main.HyperlinkRelationships);
                Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc));
            }
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Remove_on_a_layout_without_a_link_changes_nothing()
    {
        var path = CopyOfCorpus(Corpus.StandardStatement);
        try
        {
            var before = File.ReadAllBytes(path);
            using (var doc = WordprocessingDocument.Open(path, true, new OpenSettings { AutoSave = false }))
            {
                Assert.Empty(AttachedTemplateLink.Remove(doc));
            }

            Assert.Equal(before, File.ReadAllBytes(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_long_target_is_cut_where_messages_echo_it()
    {
        var described = AttachedTemplateLink.Describe(new Uri("https://example.com/" + new string('a', 300) + "?token=secret"));

        Assert.Equal(121, described.Length);
        Assert.EndsWith("…", described, StringComparison.Ordinal);
        Assert.DoesNotContain("secret", described, StringComparison.Ordinal);
    }

    [Fact]
    public void An_edit_drops_the_stock_layouts_link_and_says_so_in_its_summary()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            Assert.Equal((1, 1), Links(path));

            var response = EditTools.InsertText(path, "x", "documentEnd");

            Assert.True(response.Ok, response.Error?.Message);
            var dto = Assert.IsType<EditResultDto>(response.Data);
            Assert.Contains($"Also unlinked the external attached template '{CorpusTemplate}'", dto.Summary);
            Assert.Contains("Business Central refuses to upload a layout that links one", dto.Summary);
            Assert.True(dto.QuickValidation.Passed, "the saved layout no longer carries the attached-template error");
            Assert.Equal((0, 0), Links(path));

            // The next save has nothing left to unlink and says nothing about it.
            var again = Assert.IsType<EditResultDto>(EditTools.InsertText(path, "y", "documentEnd").Data);
            Assert.DoesNotContain("attached template", again.Summary);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_table_edit_and_a_cell_edit_say_so_too()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var table = Assert.IsType<TableEditResultDto>(
                TableTools.InsertTable(path, 1, TableArgs.Columns(2), locationType: "documentEnd").Data);
            Assert.Contains("Also unlinked the external attached template", table.Summary);

            var copy = CopyOfCorpus(Corpus.SalesInvoice);
            try
            {
                var tableIndex = Assert.IsType<TableEditResultDto>(
                    TableTools.InsertTable(copy, 1, TableArgs.Columns(2), locationType: "documentEnd").Data).TableIndex;
                LinkTemplate(copy, CorpusTemplate);

                var cell = Assert.IsType<CellEditResultDto>(EditTools.SetCellText(copy, tableIndex, 0, 0, "x").Data);
                Assert.Contains("Also unlinked the external attached template", cell.Summary);
                Assert.Equal((0, 0), Links(copy));
            }
            finally
            {
                File.Delete(copy);
            }
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Refresh_against_the_layouts_own_schema_unlinks_it_and_lists_the_target()
    {
        // The repair validate_layout's error names for a layout that needs no other edit.
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = LifecycleTools.RefreshXmlPart(path, path);

            Assert.True(response.Ok, response.Error?.Message);
            var dto = Assert.IsType<RefreshResultDto>(response.Data);
            Assert.Equal([CorpusTemplate], dto.UnlinkedAttachedTemplates);
            Assert.True(dto.QuickValidation.Passed, "the reported validation is the saved layout's, link gone");
            Assert.Equal((0, 0), Links(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Create_layout_over_a_template_that_links_one_unlinks_it_instead_of_refusing_the_template()
    {
        // A template with a BC part and zero bound controls (a blank build) whose part create_layout REPLACES. Its
        // template_not_unbound refusal keys on any quick-validation error after that, so the link has to be gone
        // before that validation runs.
        var templatePath = TempPath();
        var outputPath = TempPath();
        try
        {
            Assert.True(LifecycleTools.CreateLayout(Corpus.Path(Corpus.StandardSalesQuote), templatePath).Ok);
            LinkTemplate(templatePath, CorpusTemplate);

            var response = LifecycleTools.CreateLayout(Corpus.Path(Corpus.StandardSalesQuote), outputPath, templatePath);

            Assert.True(response.Ok, response.Error?.Message);
            var dto = Assert.IsType<CreateResultDto>(response.Data);
            Assert.True(dto.ReplacedExistingBcPart);
            Assert.Equal([CorpusTemplate], dto.UnlinkedAttachedTemplates);
            Assert.True(dto.QuickValidation.Passed);
            Assert.Equal((0, 0), Links(outputPath));
            Assert.Equal((1, 1), Links(templatePath));
        }
        finally
        {
            File.Delete(templatePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void A_blank_build_reports_nothing_unlinked()
    {
        var outputPath = TempPath();
        try
        {
            var dto = Assert.IsType<CreateResultDto>(LifecycleTools.CreateLayout(Corpus.Path(Corpus.StandardSalesQuote), outputPath).Data);

            Assert.Null(dto.UnlinkedAttachedTemplates);
        }
        finally
        {
            File.Delete(outputPath);
        }
    }
}
