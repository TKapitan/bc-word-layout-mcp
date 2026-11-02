using BcWordLayout.Domain;
using BcWordLayout.McpHost;
using BcWordLayout.McpHost.Tools;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

/// <summary>
/// Paragraph styles and clearing direct formatting (GitHub issue #146): <c>set_text_format</c>'s
/// <c>paragraphStyle</c> and <c>clearDirectFormatting</c>, and <c>set_cell_text</c>'s <c>paragraphStyle</c>. A paragraph
/// style is themable formatting, gated at BC 29 (GitHub issue #152); clearing is not.
/// </summary>
[Collection("bc-version-environment-seam")]
public sealed class ParagraphStyleTests : IDisposable
{
    /// <summary>StandardSalesInvoiceVatSpec's title: DocumentTitle_Lbl, set directly in the theme heading font, 24 pt, 0070C0.</summary>
    private const int VatSpecTitleId = 573011813;

    public ParagraphStyleTests() => ToolGuards.BcVersionEnvironmentValue = null;

    public void Dispose() => ToolGuards.BcVersionEnvironmentValue = null;

    private static string InvoiceFor(TargetFolder repo, string? application)
    {
        if (application is not null)
        {
            repo.AppJson(".", application);
        }

        return repo.CopyIn(Corpus.Path(Corpus.SalesInvoiceVatSpec), "Invoice.docx");
    }

    private static SdtElement Control(WordprocessingDocument doc, int id) =>
        doc.MainDocumentPart!.Document!.Body!.Descendants<SdtElement>()
            .Single(s => s.SdtProperties!.GetFirstChild<SdtId>()?.Val?.Value == id);

    [Fact]
    public void One_call_turns_the_plain_invoices_direct_formatted_title_into_the_themable_captures_heading()
    {
        // StandardSalesInvoiceBlueSimpleThemable styles this exact line Heading1 with an accent1 colour and no direct
        // font or size; its plain twin sets all three directly. Clearing, then styling and colouring, is the move.
        using var repo = new TargetFolder();
        var path = InvoiceFor(repo, "29.0.0.0");

        var response = EditTools.SetTextFormat(
            path, "control", controlId: VatSpecTitleId, format: new TextFormatArgument { Color = "accent1" },
            paragraphStyle: "Heading1", clearDirectFormatting: true);

        Assert.True(response.Ok, response.Error?.Message);
        var summary = Assert.IsType<EditResultDto>(response.Data).Summary;
        Assert.Contains("cleared the direct formatting (", summary, StringComparison.Ordinal);
        Assert.Contains(", then colour theme colour accent1", summary, StringComparison.Ordinal);
        Assert.Contains("paragraph style 'Heading1' (1 paragraph(s))", summary, StringComparison.Ordinal);
        Assert.DoesNotContain("which wins over the paragraph style", summary, StringComparison.Ordinal);

        using var doc = WordprocessingDocument.Open(path, false);
        var control = Control(doc, VatSpecTitleId);
        var paragraph = control.Ancestors<Paragraph>().First();
        Assert.Equal("Heading1", paragraph.ParagraphProperties!.ParagraphStyleId!.Val!.Value);
        var formatted = control.Descendants<Run>().Select(r => r.RunProperties!)
            .Append(control.SdtProperties!.GetFirstChild<RunProperties>()!);
        Assert.All(formatted, rPr =>
        {
            var only = Assert.Single(rPr.ChildElements);
            var color = Assert.IsType<Color>(only);
            Assert.Equal("accent1", color.ThemeColor!.InnerText);
        });

        // The line's own mark (24 pt) would keep the line that tall under the style; it is the styled paragraph's.
        Assert.Null(paragraph.ParagraphProperties.ParagraphMarkRunProperties);
        Assert.True(LayoutValidator.Quick(path).Findings.All(f => f.Check != "paragraph-style-resolves"));
    }

    [Fact]
    public void A_style_left_under_direct_formatting_is_reported_and_clearing_alone_is_never_gated()
    {
        using var repo = new TargetFolder();
        var path = InvoiceFor(repo, "29.0.0.0");

        var styled = EditTools.SetTextFormat(path, "control", controlId: VatSpecTitleId, paragraphStyle: "Heading1");

        Assert.True(styled.Ok, styled.Error?.Message);
        Assert.Contains(
            "2 run(s) in those paragraphs carry direct formatting (a font, size, colour and the like), which wins over "
            + "the paragraph style 'Heading1'. Pass clearDirectFormatting true",
            Assert.IsType<EditResultDto>(styled.Data).Summary,
            StringComparison.Ordinal);

        File.WriteAllText(Path.Combine(repo.Root, "app.json"), """{ "application": "28.0.0.0" }""");
        var cleared = EditTools.SetTextFormat(path, "control", controlId: VatSpecTitleId, clearDirectFormatting: true);

        Assert.True(cleared.Ok, cleared.Error?.Message);
        using var doc = WordprocessingDocument.Open(path, false);
        Assert.All(Control(doc, VatSpecTitleId).Descendants<Run>(), run => Assert.Null(run.RunProperties));
    }

    [Fact]
    public void Clearing_keeps_the_character_style_and_what_is_not_formatting()
    {
        using var repo = new TargetFolder();
        var path = InvoiceFor(repo, null);
        Assert.True(EditTools.InsertText(path, "Kept apart", "documentEnd",
            format: new TextFormatArgument { CharacterStyle = "Strong", Bold = false, FontSizePoints = 20 }).Ok);
        using (var doc = WordprocessingDocument.Open(path, true))
        {
            var run = doc.MainDocumentPart!.Document!.Body!.Descendants<Run>().Last(r => r.InnerText == "Kept apart");
            run.RunProperties!.AddChild(new Languages { Val = "da-DK" });
            run.RunProperties.AddChild(new NoProof());
            doc.MainDocumentPart.Document.Save();
        }

        var response = EditTools.SetTextFormat(path, "text", searchText: "Kept apart", clearDirectFormatting: true);

        Assert.True(response.Ok, response.Error?.Message);
        using var reopened = WordprocessingDocument.Open(path, false);
        var rPr = reopened.MainDocumentPart!.Document!.Body!.Descendants<Run>().Last(r => r.InnerText == "Kept apart").RunProperties!;
        Assert.Equal(["rStyle", "noProof", "lang"], rPr.ChildElements.Select(c => c.LocalName));
    }

    [Fact]
    public void set_cell_text_sets_the_paragraph_style_of_the_cell_it_writes()
    {
        using var repo = new TargetFolder();
        var path = InvoiceFor(repo, "29.0.0.0");
        var tableIndex = Assert.IsType<TableEditResultDto>(TableTools.InsertTable(path, 1, TableArgs.Columns(2), "documentEnd").Data).TableIndex;

        var response = EditTools.SetCellText(path, tableIndex, 0, 0, "Invoice", paragraphStyle: "Title");

        Assert.True(response.Ok, response.Error?.Message);
        Assert.Contains("Its paragraph takes the paragraph style 'Title'.", Assert.IsType<CellEditResultDto>(response.Data).Summary,
            StringComparison.Ordinal);
        using var doc = WordprocessingDocument.Open(path, false);
        var cell = TableGridNavigator.Tables(doc.MainDocumentPart!.Document!.Body!)[tableIndex].Descendants<TableCell>().First();
        Assert.Equal("Title", cell.Elements<Paragraph>().Single().ParagraphProperties!.ParagraphStyleId!.Val!.Value);
    }

    [Theory]
    [InlineData("NoSuchStyle", "the layout's styles part does not define it")]
    [InlineData("heading1", "the layout's styles part does not define it")]
    [InlineData("Strong", "it is a character style, not a paragraph style")]
    public void A_style_the_layout_does_not_define_as_a_paragraph_style_is_refused_naming_the_ones_it_does(string styleId, string reason)
    {
        using var repo = new TargetFolder();
        var path = InvoiceFor(repo, "29.0.0.0");
        var before = File.ReadAllBytes(path);

        var response = EditTools.SetTextFormat(path, "control", controlId: VatSpecTitleId, paragraphStyle: styleId);

        Assert.False(response.Ok);
        Assert.Equal(ErrorCodes.InvalidArgument, response.Error!.Code);
        Assert.Contains(reason, response.Error.Message, StringComparison.Ordinal);
        Assert.Contains("Paragraph styles this layout defines: ", response.Error.Message, StringComparison.Ordinal);
        Assert.Contains("'Heading1'", response.Error.Message, StringComparison.Ordinal);
        Assert.Contains("paragraphStyle is the id of a PARAGRAPH style", response.Error.Hint, StringComparison.Ordinal);
        Assert.Equal(before, File.ReadAllBytes(path));
    }

    [Theory]
    [InlineData("28.0.0.0")]
    [InlineData(null)]
    public void A_paragraph_style_is_refused_below_BC_29_or_for_an_unknown_target(string? application)
    {
        using var repo = new TargetFolder();
        var path = InvoiceFor(repo, application);
        var before = File.ReadAllBytes(path);

        var responses = new[]
        {
            EditTools.SetTextFormat(path, "control", controlId: VatSpecTitleId, paragraphStyle: "Heading1"),
            EditTools.SetCellText(path, 0, 0, 0, "x", paragraphStyle: "Heading1"),
        };

        Assert.All(responses, r => Assert.Equal(ErrorCodes.BcVersionUnsupported, r.Error?.Code));
        Assert.Equal(before, File.ReadAllBytes(path));
    }
}
