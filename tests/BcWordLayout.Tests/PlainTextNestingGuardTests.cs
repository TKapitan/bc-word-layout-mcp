using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

/// <summary>
/// Covers <see cref="PlainTextNestingGuard"/> — the detector for content controls nested inside a
/// plain-text content control (a combination Word rejects as corrupt but OpenXmlValidator accepts). Every
/// clean corpus layout must report zero; a layout with a control anchored inside a plain-text control must
/// report it.
/// </summary>
public class PlainTextNestingGuardTests
{
    [Theory]
    [InlineData(Corpus.SalesInvoice)]
    [InlineData(Corpus.InventoryOrderDetails)]
    [InlineData(Corpus.StandardStatement)]
    public void Find_reports_nothing_for_a_clean_corpus_layout(string corpusFile)
    {
        using var doc = WordprocessingDocument.Open(Corpus.Path(corpusFile), false);
        Assert.Empty(PlainTextNestingGuard.Find(doc));
    }

    [Fact]
    public void Find_reports_a_field_anchored_inside_a_cell_level_plaintext_control()
    {
        // YourReference_Lbl (id -1130623254) is a cell-level plain-text control; inserting a field
        // afterControl it anchors the new field inside that control's cell (see LayoutEditor).
        const int cellLevelPlainTextLabelId = -1130623254;
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-ptguard-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(Corpus.SalesInvoice), path, overwrite: true);
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                LayoutEditor.InsertField(
                    doc,
                    "/Header/SalesPersonName",
                    new Location { Type = LocationKind.AfterControl, ControlId = cellLevelPlainTextLabelId });
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var nestings = PlainTextNestingGuard.Find(reopened);

            var offender = Assert.Single(nestings);
            Assert.Equal(cellLevelPlainTextLabelId, offender.OuterId);
            Assert.Equal("document.xml", offender.Part);
            Assert.Contains("plain-text", offender.Describe(), StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
