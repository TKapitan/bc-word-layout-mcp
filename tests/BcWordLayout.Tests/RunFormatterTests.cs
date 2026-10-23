using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

/// <summary>
/// Covers <see cref="RunFormatter"/>, the one place a <see cref="TextFormat"/> is checked and written (GitHub
/// issue #15): the OOXML each member writes, the explicit-off rule over a character style, the refusals, and
/// that the same code serves a run, a control's <c>w:sdtPr/w:rPr</c> and a paragraph mark.
/// </summary>
public class RunFormatterTests
{
    // StandardSalesQuote header2: DocumentNo's runs carry w:rStyle 'Strong' (see InheritedRunFormatTests).
    private const int DocumentNoId = 1543556682;

    private static string CopyOfCorpus(string corpusFile)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-runformat-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    private static RunFormatter.Prepared Prepare(WordprocessingDocument doc, TextFormat format) =>
        RunFormatter.Prepare(format, doc) ?? throw new InvalidOperationException("Nothing to apply.");

    /// <summary>Applies <paramref name="format"/> to a fresh run's properties, against a corpus copy's styles.</summary>
    private static RunProperties Formatted(TextFormat format, RunProperties? start = null, string corpus = Corpus.StandardSalesQuote)
    {
        var path = CopyOfCorpus(corpus);
        try
        {
            using var doc = WordprocessingDocument.Open(path, false);
            var rPr = start ?? new RunProperties();
            RunFormatter.Apply(rPr, Prepare(doc, format));
            return rPr;
        }
        finally
        {
            File.Delete(path);
        }
    }

    private static ArgumentException Refusal(TextFormat format, string corpus = Corpus.StandardSalesQuote)
    {
        var path = CopyOfCorpus(corpus);
        try
        {
            using var doc = WordprocessingDocument.Open(path, false);
            return Assert.Throws<ArgumentException>(() => RunFormatter.Prepare(format, doc));
        }
        finally
        {
            File.Delete(path);
        }
    }

    private static List<ValidationErrorInfo> SchemaErrors(OpenXmlElement element)
    {
        var doc = new Document(new Body(new Paragraph(element is RunProperties rPr
            ? new Run(rPr, new Text("x"))
            : element)));
        return new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc).ToList();
    }

    // ---- what each member writes ----

    [Fact]
    public void Each_member_writes_its_corpus_shape_in_schema_order()
    {
        var rPr = Formatted(new TextFormat
        {
            Color = "0070c0",
            FontName = "Segoe UI",
            FontSizePoints = 8,
            Underline = true,
            Italic = true,
            Bold = true,
            CharacterStyle = "Strong",
        });

        Assert.Equal("Strong", rPr.RunStyle?.Val?.Value);
        Assert.NotNull(rPr.Bold);
        Assert.Null(rPr.Bold!.Val);
        Assert.NotNull(rPr.Italic);
        Assert.NotNull(rPr.ItalicComplexScript); // w:i always comes with w:iCs in the corpus styles
        Assert.Equal(UnderlineValues.Single, rPr.Underline?.Val?.Value);
        Assert.Equal("16", rPr.FontSize?.Val?.Value);
        Assert.Equal("16", rPr.FontSizeComplexScript?.Val?.Value);
        Assert.Equal("Segoe UI", rPr.RunFonts?.Ascii?.Value);
        Assert.Equal("Segoe UI", rPr.RunFonts?.HighAnsi?.Value);
        Assert.Equal("Segoe UI", rPr.RunFonts?.ComplexScript?.Value);
        Assert.Equal("0070C0", rPr.Color?.Val?.Value); // normalized to upper case, as the corpus writes it

        Assert.Empty(SchemaErrors(rPr));
    }

    [Theory]
    [InlineData("#1f4e79", "1F4E79")]
    [InlineData(" 0070C0 ", "0070C0")]
    [InlineData("AUTO", "auto")]
    public void Color_accepts_hex_with_or_without_a_hash_and_auto(string given, string written)
    {
        Assert.Equal(written, Formatted(new TextFormat { Color = given }).Color?.Val?.Value);
    }

    [Fact]
    public void Font_and_colour_drop_the_theme_attributes_that_would_win_over_them_and_keep_the_east_asian_font()
    {
        // Word lets asciiTheme/hAnsiTheme/cstheme and themeColor win over the explicit values, so keeping them
        // would make the call do nothing visible. The corpus has 'cstheme=minorHAnsi' runs and accent1 colours.
        var start = new RunProperties(
            new RunFonts
            {
                AsciiTheme = ThemeFontValues.MinorHighAnsi,
                HighAnsiTheme = ThemeFontValues.MinorHighAnsi,
                ComplexScriptTheme = ThemeFontValues.MinorHighAnsi,
                EastAsiaTheme = ThemeFontValues.MinorEastAsia,
            },
            new Color { Val = "1F4E79", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "80" });

        var rPr = Formatted(new TextFormat { FontName = "Arial", Color = "auto" }, start);

        Assert.Equal("Arial", rPr.RunFonts!.Ascii?.Value);
        Assert.Null(rPr.RunFonts.AsciiTheme);
        Assert.Null(rPr.RunFonts.HighAnsiTheme);
        Assert.Null(rPr.RunFonts.ComplexScriptTheme);
        Assert.Equal(ThemeFontValues.MinorEastAsia, rPr.RunFonts.EastAsiaTheme?.Value);
        Assert.Equal("auto", rPr.Color!.Val?.Value);
        Assert.Null(rPr.Color.ThemeColor);
        Assert.Null(rPr.Color.ThemeShade);
    }

    [Fact]
    public void Members_left_null_are_not_touched()
    {
        var start = new RunProperties(new Bold(), new Underline { Val = UnderlineValues.Single }, new FontSize { Val = "18" });

        var rPr = Formatted(new TextFormat { Italic = true }, start);

        Assert.NotNull(rPr.Bold);
        Assert.Equal(UnderlineValues.Single, rPr.Underline?.Val?.Value);
        Assert.Equal("18", rPr.FontSize?.Val?.Value);
        Assert.NotNull(rPr.Italic);
    }

    [Fact]
    public void An_empty_format_prepares_to_nothing()
    {
        var path = CopyOfCorpus(Corpus.StandardSalesQuote);
        try
        {
            using var doc = WordprocessingDocument.Open(path, false);
            Assert.Null(RunFormatter.Prepare(new TextFormat(), doc));
            Assert.Null(RunFormatter.Prepare(null, doc));
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- false: removal, or an explicit off over a character style ----

    [Fact]
    public void False_removes_the_property_when_no_character_style_could_supply_it()
    {
        var start = new RunProperties(new Bold(), new Italic(), new ItalicComplexScript(), new Underline { Val = UnderlineValues.Single });

        var rPr = Formatted(new TextFormat { Bold = false, Italic = false, Underline = false }, start);

        Assert.Null(rPr.Bold);
        Assert.Null(rPr.Italic);
        Assert.Null(rPr.ItalicComplexScript);
        Assert.Null(rPr.Underline);
    }

    [Fact]
    public void False_writes_an_explicit_off_over_a_character_style()
    {
        // Removing w:b cannot un-bold 'Strong' (GitHub issue #122); the same holds for italic and underline
        // that a style supplies (the corpus styles include Emphasis and underlined ones).
        var start = new RunProperties(new RunStyle { Val = "Strong" });

        var rPr = Formatted(new TextFormat { Bold = false, Italic = false, Underline = false }, start);

        Assert.False(rPr.Bold?.Val?.Value ?? true);
        Assert.False(rPr.Italic?.Val?.Value ?? true);
        Assert.Equal(UnderlineValues.None, rPr.Underline?.Val?.Value);
        Assert.Empty(SchemaErrors(rPr));
    }

    // ---- the same code on a paragraph mark ----

    [Fact]
    public void Apply_serves_a_paragraph_mark_too()
    {
        var mark = new ParagraphMarkRunProperties(new FontSize { Val = "18" });

        var path = CopyOfCorpus(Corpus.StandardSalesQuote);
        try
        {
            using var doc = WordprocessingDocument.Open(path, false);
            RunFormatter.Apply(mark, Prepare(doc, new TextFormat { Bold = true, Italic = true, FontName = "Segoe UI", CharacterStyle = "Strong" }));
        }
        finally
        {
            File.Delete(path);
        }

        Assert.NotNull(mark.GetFirstChild<Bold>());
        Assert.NotNull(mark.GetFirstChild<Italic>());
        Assert.Equal("Segoe UI", mark.GetFirstChild<RunFonts>()?.Ascii?.Value);
        Assert.Equal("Strong", mark.GetFirstChild<RunStyle>()?.Val?.Value);
        Assert.Empty(SchemaErrors(new ParagraphProperties(mark)));
    }

    // ---- refusals ----

    [Theory]
    [InlineData(3.5)]
    [InlineData(96.5)]
    public void A_size_outside_4_to_96_is_refused(double points)
    {
        var ex = Refusal(new TextFormat { FontSizePoints = points });
        Assert.Equal("format", ex.ParamName);
        Assert.Contains("format.fontSizePoints must be between 4 and 96", ex.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("blue")]
    [InlineData("0070C")]
    [InlineData("#GG0000")]
    public void A_colour_that_is_not_hex_or_auto_is_refused(string color)
    {
        var ex = Refusal(new TextFormat { Color = color });
        Assert.Equal("format", ex.ParamName);
        Assert.Contains("format.color must be a colour as six hex digits", ex.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("   ")]
    [InlineData("Segoe\tUI")]
    public void A_blank_or_control_character_font_name_is_refused(string font)
    {
        var ex = Refusal(new TextFormat { FontName = font });
        Assert.Equal("format", ex.ParamName);
        Assert.Contains("format.fontName must be a font family name", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_character_style_the_layout_does_not_define_is_refused_naming_the_ones_it_does()
    {
        var ex = Refusal(new TextFormat { CharacterStyle = "strong" }); // ids are case-sensitive

        Assert.Equal("format", ex.ParamName);
        Assert.Contains("format.characterStyle 'strong' cannot be applied: the layout's styles part does not define it", ex.Message, StringComparison.Ordinal);
        Assert.Contains("'Strong'", ex.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("'DefaultParagraphFont'", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_paragraph_style_is_refused_as_a_character_style()
    {
        // Heading1 is a paragraph style; its linked character style is Heading1Char.
        var ex = Refusal(new TextFormat { CharacterStyle = "Heading1" });

        Assert.Contains("it is a paragraph style, not a character style", ex.Message, StringComparison.Ordinal);
        Assert.Contains("'Heading1Char'", ex.Message, StringComparison.Ordinal);
    }

    // ---- end to end through an insert ----

    [Fact]
    public void An_insert_with_every_member_validates_cleanly_on_a_stock_layout()
    {
        var path = CopyOfCorpus(Corpus.StandardSalesQuote);
        try
        {
            int errorsBefore;
            using (var before = WordprocessingDocument.Open(path, false))
            {
                errorsBefore = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(before).Count();
            }

            EditResult result;
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                result = LayoutEditor.InsertField(
                    doc,
                    "/Header/DocumentDate",
                    new Location { Type = LocationKind.AfterControl, ControlId = DocumentNoId, Part = LayoutPart.Header, PartName = "header2.xml" },
                    new TextFormat { Italic = true, Underline = true, FontName = "Segoe UI", Color = "0070C0", Bold = false });
                doc.MainDocumentPart!.HeaderParts.ToList().ForEach(h => h.Header!.Save());
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var control = reopened.MainDocumentPart!.HeaderParts
                .SelectMany(h => h.Header!.Descendants<SdtRun>())
                .Single(s => SdtInspector.ReadControlId(s) == result.ControlId);

            // Over the 'Strong' it inherits, on the control's own rPr and its content alike.
            foreach (var rPr in control.Descendants<RunProperties>())
            {
                Assert.Equal("Strong", rPr.RunStyle?.Val?.Value);
                Assert.False(rPr.Bold?.Val?.Value ?? true);
                Assert.NotNull(rPr.Italic);
                Assert.Equal(UnderlineValues.Single, rPr.Underline?.Val?.Value);
                Assert.Equal("Segoe UI", rPr.RunFonts?.Ascii?.Value);
                Assert.Equal("0070C0", rPr.Color?.Val?.Value);
            }

            Assert.Equal(errorsBefore, new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened).Count());
        }
        finally
        {
            File.Delete(path);
        }
    }
}
