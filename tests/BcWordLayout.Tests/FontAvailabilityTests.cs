using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

/// <summary>
/// Fonts Business Central online lacks (<see cref="BcOnlineFonts"/>): sandbox probe P13 (BC 28) printed a line set
/// in a font that exists nowhere in a serif face. <c>validate_layout</c> warns (<c>font-installed</c>), and a
/// tool given such a <c>format.fontName</c> says so in its summary.
/// </summary>
public class FontAvailabilityTests
{
    private const string MissingFont = "BCWL Missing Font";

    private static Location DocumentEnd => new() { Type = LocationKind.DocumentEnd };

    [Theory]
    [InlineData("Segoe UI")]
    [InlineData("segoe ui semibold")]
    [InlineData("Georgia")]
    [InlineData("Aptos")]
    [InlineData("IDAutomationHC39M")]
    [InlineData("IDAutomationC11L")]
    [InlineData("IDAutomationC11L.ttf")]
    [InlineData("Bahnschrift SemiBold SemiCondensed")] // listed cut to 31 characters
    [InlineData("Neue Haas Grotesk Text Pro Black")]
    [InlineData(" Times New Roman ")]
    public void Microsofts_listed_fonts_are_installed(string font) =>
        Assert.True(BcOnlineFonts.IsInstalled(font));

    [Theory]
    [InlineData(MissingFont)]
    [InlineData("Helvetica")]
    [InlineData("Segoe")]
    [InlineData("Bahnschrift SemiBold SemiConde")] // shorter than the cut name it resembles
    public void Other_fonts_are_not(string font) =>
        Assert.False(BcOnlineFonts.IsInstalled(font));

    [Fact]
    public void No_corpus_layout_names_a_font_BC_online_lacks()
    {
        foreach (var file in Directory.GetFiles(Path.GetDirectoryName(Corpus.Path(Corpus.StandardSalesQuote))!, "*.docx"))
        {
            using var doc = WordprocessingDocument.Open(file, false);
            Assert.DoesNotContain(LayoutValidator.Quick(doc).Findings, f => f.Check == "font-installed");
        }
    }

    [Fact]
    public void A_run_in_a_missing_font_is_one_warning_per_font_naming_its_parts()
    {
        var path = Corpus.UnlinkedCopy(Corpus.StandardSalesQuote);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var format = new TextFormat { FontName = MissingFont };
            var first = LayoutEditor.InsertText(doc, "one", DocumentEnd, format);
            LayoutEditor.InsertText(doc, "two", DocumentEnd, format);
            LayoutEditor.InsertText(doc, "three", new Location { Type = LocationKind.DocumentEnd, Part = LayoutPart.Header }, format);

            var finding = Assert.Single(LayoutValidator.Quick(doc).Findings, f => f.Check == "font-installed");
            Assert.Equal(FindingSeverity.Warning, finding.Severity);
            Assert.Contains($"'{MissingFont}'", finding.Message, StringComparison.Ordinal);
            Assert.StartsWith("document.xml, header", finding.Location, StringComparison.Ordinal);

            Assert.Contains($"'{MissingFont}' is not a font Business Central online has installed", first.Summary, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void An_installed_font_gets_no_note_and_no_warning()
    {
        var path = Corpus.UnlinkedCopy(Corpus.StandardSalesQuote);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var result = LayoutEditor.InsertText(doc, "Georgia", DocumentEnd, new TextFormat { FontName = "Georgia" });

            Assert.DoesNotContain("Business Central online has installed", result.Summary, StringComparison.Ordinal);
            Assert.DoesNotContain(LayoutValidator.Quick(doc).Findings, f => f.Check == "font-installed");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_font_the_layout_embeds_is_not_reported()
    {
        var path = Corpus.UnlinkedCopy(Corpus.StandardSalesQuote);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var main = doc.MainDocumentPart!;
            var fontTable = main.FontTablePart ?? main.AddNewPart<FontTablePart>();
            fontTable.Fonts ??= new Fonts();
            var fontPart = fontTable.AddFontPart(FontPartType.FontOdttf);
            using (var data = new MemoryStream(new byte[16]))
            {
                fontPart.FeedData(data);
            }

            fontTable.Fonts.AppendChild(new Font(
                new EmbedRegularFont { Id = fontTable.GetIdOfPart(fontPart), FontKey = "{00000000-0000-0000-0000-000000000000}" })
            {
                Name = MissingFont,
            });

            var result = LayoutEditor.InsertText(doc, "embedded", DocumentEnd, new TextFormat { FontName = MissingFont });

            Assert.DoesNotContain("Business Central online has installed", result.Summary, StringComparison.Ordinal);
            Assert.DoesNotContain(LayoutValidator.Quick(doc).Findings, f => f.Check == "font-installed");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_theme_font_counts_only_where_something_refers_to_it()
    {
        var path = Corpus.UnlinkedCopy(Corpus.StandardSalesQuote);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var main = doc.MainDocumentPart!;
            var scheme = main.ThemePart!.Theme!.ThemeElements!.FontScheme!;
            scheme.MajorFont!.LatinFont!.Typeface = MissingFont;
            scheme.MinorFont!.LatinFont!.Typeface = MissingFont + " Body";

            // The minor (body) font is reached through the theme references every Word default style carries;
            // nothing in the quote asks for the major (heading) font.
            var usesMajor = main.StyleDefinitionsPart!.Styles!.Descendants<RunFonts>().Concat(main.Document!.Descendants<RunFonts>())
                .Any(f => f.AsciiTheme?.Value == ThemeFontValues.MajorHighAnsi || f.HighAnsiTheme?.Value == ThemeFontValues.MajorHighAnsi
                    || f.AsciiTheme?.Value == ThemeFontValues.MajorAscii || f.HighAnsiTheme?.Value == ThemeFontValues.MajorAscii);

            var findings = LayoutValidator.Quick(doc).Findings.Where(f => f.Check == "font-installed").ToList();
            var body = Assert.Single(findings, f => f.Message.Contains($"'{MissingFont} Body'", StringComparison.Ordinal));
            Assert.StartsWith("theme", body.Location, StringComparison.Ordinal);
            Assert.Equal(usesMajor, findings.Any(f => f.Message.Contains($"'{MissingFont}' ", StringComparison.Ordinal)));
        }
        finally
        {
            File.Delete(path);
        }
    }
}
