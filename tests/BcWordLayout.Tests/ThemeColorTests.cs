using System.Globalization;
using BcWordLayout.Domain;
using BcWordLayout.McpHost;
using BcWordLayout.McpHost.Tools;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

/// <summary>
/// Theme colours (GitHub issue #145): <see cref="ThemePalette"/> resolving a slot as Word caches it, <see cref="LayoutColor"/>
/// parsing every colour argument, and the tools writing <c>w:themeColor</c> / <c>w:themeFill</c> behind the
/// <c>themable-formatting</c> gate (BC 29, GitHub issue #152).
/// </summary>
[Collection("bc-version-environment-seam")]
public sealed class ThemeColorTests : IDisposable
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    public ThemeColorTests() => ToolGuards.BcVersionEnvironmentValue = null;

    public void Dispose() => ToolGuards.BcVersionEnvironmentValue = null;

    private static string[] CorpusFiles() =>
        Directory.GetFiles(Path.Combine(AppContext.BaseDirectory, "corpus"), "*.docx");

    /// <summary>Every theme-coloured element in the corpus's content and styles parts: slot, tint, shade, cached value.</summary>
    private static IEnumerable<(string File, string Slot, string? Tint, string? Shade, string Cached, string Resolved)> CorpusThemeColours()
    {
        foreach (var file in CorpusFiles())
        {
            using var doc = WordprocessingDocument.Open(file, false);
            var main = doc.MainDocumentPart!;
            var roots = PartWalker.ContentParts(main).Select(p => (OpenXmlElement)p.Root)
                .Append(main.StyleDefinitionsPart!.Styles!);
            foreach (var element in roots.SelectMany(r => r.Descendants()))
            {
                foreach (var (theme, cached, tint, shade) in new[]
                         {
                             ("themeColor", element is Color ? "val" : "color", "themeTint", "themeShade"),
                             ("themeFill", "fill", "themeFillTint", "themeFillShade"),
                         })
                {
                    var slot = Attr(element, theme);
                    if (slot is null || !ThemePalette.Slots.Contains(slot))
                    {
                        continue;
                    }

                    yield return (Path.GetFileName(file), slot, Attr(element, tint), Attr(element, shade), Attr(element, cached)!,
                        ThemePalette.Resolve(main, slot)!);
                }
            }
        }
    }

    private static string? Attr(OpenXmlElement element, string localName) =>
        element.GetAttributes().FirstOrDefault(a => a.NamespaceUri == W && a.LocalName == localName).Value;

    // ---- ThemePalette: what Word caches ----

    [Fact]
    public void Every_untinted_theme_colour_in_the_corpus_caches_the_colour_its_slot_resolves_to()
    {
        var plain = CorpusThemeColours().Where(c => c.Tint is null && c.Shade is null).ToList();

        Assert.True(plain.Count > 200, $"only {plain.Count} untinted theme colours found");
        Assert.All(plain, c => Assert.True(c.Cached == c.Resolved, $"{c.File}: {c.Slot} caches {c.Cached}, resolves to {c.Resolved}"));
    }

    [Fact]
    public void Tinted_and_shaded_theme_colours_match_the_corpus_cache_to_within_one_step()
    {
        // Word's arithmetic is undocumented: truncating the float HSL result reproduces 14 of the 18 distinct cached
        // values exactly, the rules a tool would redraw among them, and the other four within one step per channel.
        var distinct = CorpusThemeColours()
            .Where(c => c.Tint is not null || c.Shade is not null)
            .Select(c => (c.Resolved, c.Tint, c.Shade, c.Cached))
            .Distinct()
            .ToList();

        Assert.Equal(18, distinct.Count);
        var exact = 0;
        foreach (var (rgb, tint, shade, cached) in distinct)
        {
            var computed = ThemePalette.ApplyTintShade(rgb, tint, shade);
            for (var i = 0; i < 6; i += 2)
            {
                var difference = int.Parse(computed.AsSpan(i, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture)
                                 - int.Parse(cached.AsSpan(i, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                Assert.True(Math.Abs(difference) <= 1, $"{rgb} tint {tint} shade {shade}: computed {computed}, Word cached {cached}");
            }

            exact += computed == cached ? 1 : 0;
        }

        Assert.Equal(14, exact);
        Assert.Equal("9B9482", ThemePalette.ApplyTintShade("4C483D", "99", null)); // text2 lighter 40% in VatSpec's rules
        Assert.Equal("AEAAAA", ThemePalette.ApplyTintShade("E7E6E6", null, "BF")); // background2 darker 25%
    }

    [Fact]
    public void A_slot_resolves_through_the_colour_scheme_mapping()
    {
        using var doc = WordprocessingDocument.Open(Corpus.Path(Corpus.SalesInvoiceBlueSimpleThemable), false);
        var main = doc.MainDocumentPart!;

        Assert.Equal("F24F4F", ThemePalette.Resolve(main, "accent1"));
        Assert.Equal("000000", ThemePalette.Resolve(main, "text1")); // dk1 is the system colour windowText
        Assert.Equal("4C483D", ThemePalette.Resolve(main, "text2")); // dk2
        Assert.Equal("E4E3E2", ThemePalette.Resolve(main, "background2")); // lt2
    }

    // ---- LayoutColor: the argument ----

    [Theory]
    [InlineData("0070c0", "0070C0", null, null, null)]
    [InlineData("#E8E8E8", "E8E8E8", null, null, null)]
    [InlineData("AUTO", "auto", null, null, null)]
    [InlineData("accent1", "F24F4F", "accent1", null, null)]
    [InlineData(" Accent1 ", "F24F4F", "accent1", null, null)]
    [InlineData("text2 lighter 40%", "9B9482", "text2", "99", null)]
    [InlineData("text2 LIGHTER 40", "9B9482", "text2", "99", null)]
    [InlineData("accent1 lighter 80%", null, "accent1", "33", null)]
    [InlineData("background1 darker 5%", "F2F2F2", "background1", null, "F2")]
    [InlineData("background1 darker 25%", "BFBFBF", "background1", null, "BF")]
    [InlineData("text1 lighter 15%", "262626", "text1", "D9", null)]
    [InlineData("text1 lighter 35%", "595959", "text1", "A6", null)]
    [InlineData("followedhyperlink", null, "followedHyperlink", null, null)]
    public void A_colour_argument_is_a_literal_or_a_theme_slot_with_Words_palette_percentages(
        string argument, string? value, string? slot, string? tint, string? shade)
    {
        using var doc = WordprocessingDocument.Open(Corpus.Path(Corpus.SalesInvoiceBlueSimpleThemable), false);

        var color = LayoutColor.Parse(argument, "format.color", "format", doc, ColorUse.Text);

        if (value is not null)
        {
            Assert.Equal(value, color.Value);
        }

        Assert.Equal(slot, color.ThemeSlot);
        Assert.Equal(tint, color.Tint);
        Assert.Equal(shade, color.Shade);
        Assert.Equal(slot is not null, LayoutColor.NamesThemeSlot(argument));
    }

    [Theory]
    [InlineData("blue", "must be six hex digits RRGGBB")]
    [InlineData("accent7", "must be six hex digits RRGGBB")]
    [InlineData("0070C", "must be six hex digits RRGGBB")]
    [InlineData("accent1 lighter 0%", "the percentage must be between 1 and 99")]
    [InlineData("accent1 brighter 40%", "must be six hex digits RRGGBB")]
    public void A_malformed_colour_is_refused_naming_every_accepted_form(string argument, string expected)
    {
        using var doc = WordprocessingDocument.Open(Corpus.Path(Corpus.SalesInvoiceBlueSimpleThemable), false);

        var ex = Assert.Throws<ArgumentException>(() => LayoutColor.Parse(argument, "format.color", "format", doc, ColorUse.Text));

        Assert.Contains(expected, ex.Message, StringComparison.Ordinal);
        Assert.Equal("format", ex.ParamName);
    }

    [Fact]
    public void A_lighter_fill_is_refused_because_no_captured_layout_carries_one()
    {
        using var doc = WordprocessingDocument.Open(Corpus.Path(Corpus.SalesInvoiceBlueSimpleThemable), false);

        var ex = Assert.Throws<ArgumentException>(() => LayoutColor.Parse("accent1 lighter 80%", "fill", "fill", doc, ColorUse.Fill));

        Assert.Contains("w:themeFillTint", ex.Message, StringComparison.Ordinal);
        Assert.Equal("BF", LayoutColor.Parse("accent1 darker 25%", "fill", "fill", doc, ColorUse.Fill).Shade);
        Assert.Equal("33", LayoutColor.Parse("accent1 lighter 80%", "borderColor", "borderColor", doc, ColorUse.Rule).Tint);
    }

    [Fact]
    public void A_theme_slot_is_refused_in_a_layout_without_a_theme_part()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.PlainParagraph("x"));
        try
        {
            using var doc = WordprocessingDocument.Open(path, false);
            Assert.Null(doc.MainDocumentPart!.ThemePart);

            var ex = Assert.Throws<ArgumentException>(() => LayoutColor.Parse("accent1", "format.color", "format", doc, ColorUse.Text));

            Assert.Contains("this layout has no theme part", ex.Message, StringComparison.Ordinal);
            Assert.Equal("0070C0", LayoutColor.Parse("0070C0", "format.color", "format", doc, ColorUse.Text).Value);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- the tools ----

    [Fact]
    public void insert_text_writes_a_theme_colour_as_Word_does_for_a_BC_29_target()
    {
        using var repo = new TargetFolder();
        repo.AppJson(".", "29.0.0.0");
        var path = repo.CopyIn(Corpus.Path(Corpus.SalesInvoiceVatSpec), "Invoice.docx");

        var response = EditTools.InsertText(path, "Themed", "documentEnd", format: new TextFormatArgument { Color = "accent1" });

        Assert.True(response.Ok, response.Error?.Message);
        using var doc = WordprocessingDocument.Open(path, false);
        var color = RunWithText(doc, "Themed").RunProperties!.Color!;
        Assert.Equal("F24F4F", color.Val!.Value);
        Assert.Equal("accent1", color.ThemeColor!.InnerText);
        Assert.Null(color.ThemeTint);
        Assert.DoesNotContain(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc), e => e.Path?.XPath?.Contains("w:color") == true);
    }

    [Theory]
    [InlineData("28.0.0.0")]
    [InlineData(null)]
    public void A_theme_colour_is_refused_below_BC_29_or_for_an_unknown_target_and_the_file_is_untouched(string? application)
    {
        using var repo = new TargetFolder();
        if (application is not null)
        {
            repo.AppJson(".", application);
        }

        var path = repo.CopyIn(Corpus.Path(Corpus.SalesInvoiceVatSpec), "Invoice.docx");
        var before = File.ReadAllBytes(path);

        var responses = new[]
        {
            EditTools.InsertText(path, "Themed", "documentEnd", format: new TextFormatArgument { Color = "accent1" }),
            EditTools.SetTextFormat(path, "text", searchText: "Themed", format: new TextFormatArgument { Color = "text2 lighter 40%" }),
            TableTools.SetCellFormat(path, tableIndex: 0, row: 0, fill: "accent1"),
        };

        Assert.All(responses, r =>
        {
            Assert.False(r.Ok);
            Assert.Equal(ErrorCodes.BcVersionUnsupported, r.Error!.Code);
            Assert.Contains("'themable-formatting' needs BC 29.0", r.Error.Hint, StringComparison.Ordinal);
        });
        Assert.Equal(before, File.ReadAllBytes(path));

        // A literal colour is not themable formatting: it is never gated.
        Assert.True(EditTools.InsertText(path, "Blue", "documentEnd", format: new TextFormatArgument { Color = "0070C0" }).Ok);
    }

    [Fact]
    public void set_text_format_says_when_a_literal_colour_replaces_a_theme_colour()
    {
        using var repo = new TargetFolder();
        repo.AppJson(".", "28.0.0.0"); // a literal colour needs no BC 29
        var path = repo.CopyIn(Corpus.Path(Corpus.SalesInvoiceBlueSimpleThemable), "Invoice.docx");
        int controlId;
        using (var doc = WordprocessingDocument.Open(path, false))
        {
            controlId = doc.MainDocumentPart!.Document!.Body!.Descendants<SdtRun>()
                .First(s => s.SdtProperties!.GetFirstChild<RunProperties>()?.Color?.ThemeColor is not null)
                .SdtProperties!.GetFirstChild<SdtId>()!.Val!.Value;
        }

        var response = EditTools.SetTextFormat(path, "control", controlId: controlId, format: new TextFormatArgument { Color = "0070C0" });

        Assert.True(response.Ok, response.Error?.Message);
        Assert.Contains("had a theme colour, which the literal colour replaces", Assert.IsType<EditResultDto>(response.Data).Summary,
            StringComparison.Ordinal);
        using var reopened = WordprocessingDocument.Open(path, false);
        var control = reopened.MainDocumentPart!.Document!.Body!.Descendants<SdtRun>()
            .Single(s => s.SdtProperties!.GetFirstChild<SdtId>()!.Val!.Value == controlId);
        var color = control.SdtProperties!.GetFirstChild<RunProperties>()!.Color!;
        Assert.Equal("0070C0", color.Val!.Value);
        Assert.Null(color.ThemeColor);
    }

    [Fact]
    public void set_cell_format_fills_with_a_theme_colour_in_the_corpus_shape()
    {
        using var repo = new TargetFolder();
        repo.AppJson(".", "29.0.0.0");
        var path = repo.CopyIn(Corpus.Path(Corpus.SalesInvoiceVatSpec), "Invoice.docx");
        var tableIndex = Assert.IsType<TableEditResultDto>(
            TableTools.InsertTable(path, 2, TableArgs.Columns(2), "documentEnd").Data).TableIndex;

        var accent = TableTools.SetCellFormat(path, tableIndex, row: 0, fill: "accent1");
        var shaded = TableTools.SetCellFormat(path, tableIndex, row: 1, fill: "background2 darker 25%");
        var literal = TableTools.SetCellFormat(path, tableIndex, row: 0, fill: "E8E8E8");

        Assert.True(accent.Ok, accent.Error?.Message);
        Assert.True(shaded.Ok, shaded.Error?.Message);
        Assert.Contains("filled them with theme colour accent1 (F24F4F", Assert.IsType<TableEditResultDto>(accent.Data).Summary,
            StringComparison.Ordinal);
        Assert.Contains("2 fill(s) had a theme colour, which the literal colour replaces",
            Assert.IsType<TableEditResultDto>(literal.Data).Summary, StringComparison.Ordinal);

        using var doc = WordprocessingDocument.Open(path, false);
        var table = TableGridNavigator.Tables(doc.MainDocumentPart!.Document!.Body!)[tableIndex];
        var rows = table.Elements<TableRow>().ToList();
        var first = rows[0].Elements<TableCell>().First().TableCellProperties!.Shading!;
        Assert.Equal("E8E8E8", first.Fill!.Value);
        Assert.Null(first.ThemeFill);
        var second = rows[1].Elements<TableCell>().First().TableCellProperties!.Shading!;
        Assert.Equal("<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"ADAAA6\" w:themeFill=\"background2\" w:themeFillShade=\"BF\" "
                     + $"xmlns:w=\"{W}\" />", second.OuterXml);
        Assert.Equal(ThemePalette.ApplyTintShade("E4E3E2", null, "BF"), second.Fill!.Value);
    }

    private static Run RunWithText(WordprocessingDocument doc, string text) =>
        doc.MainDocumentPart!.Document!.Body!.Descendants<Run>().Last(r => r.InnerText == text);
}
