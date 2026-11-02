using System.Xml.Linq;
using BcWordLayout.Domain;
using BcWordLayout.McpHost;
using BcWordLayout.McpHost.Tools;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;

namespace BcWordLayout.Tests;

/// <summary>
/// A layout's own theme (GitHub issue #150): <c>set_layout_theme</c> changing the theme's colours and fonts and the
/// colour every theme reference caches, and <c>create_layout</c>'s <c>theme: "office"</c> giving a new layout the
/// stock Office theme. The tool tests take their BC 29 target from an <c>app.json</c> beside the layout.
/// </summary>
[Collection("bc-version-environment-seam")]
public sealed class LayoutThemeTests : IDisposable
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    public LayoutThemeTests() => ToolGuards.BcVersionEnvironmentValue = null;

    public void Dispose() => ToolGuards.BcVersionEnvironmentValue = null;

    /// <summary>Every theme reference in the content, styles and numbering parts: slot, tint, shade and cached colour.</summary>
    private static List<(string Slot, string? Tint, string? Shade, string Cached)> ThemeReferences(string path)
    {
        using var doc = WordprocessingDocument.Open(path, false);
        var main = doc.MainDocumentPart!;
        var roots = PartWalker.ContentParts(main).Select(p => p.Root).Cast<OpenXmlElement>().ToList();
        roots.Add(main.StyleDefinitionsPart!.Styles!);
        if (main.NumberingDefinitionsPart?.Numbering is { } numbering)
        {
            roots.Add(numbering);
        }

        var found = new List<(string, string?, string?, string)>();
        foreach (var element in roots.SelectMany(r => r.Descendants()))
        {
            foreach (var (theme, cached, tint, shade) in new[]
                     {
                         ("themeColor", element is Color ? "val" : "color", "themeTint", "themeShade"),
                         ("themeFill", "fill", "themeFillTint", "themeFillShade"),
                     })
            {
                if (Attr(element, theme) is { } slot && Attr(element, cached) is { } value)
                {
                    found.Add((slot, Attr(element, tint), Attr(element, shade), value));
                }
            }
        }

        return found;
    }

    private static string? Attr(OpenXmlElement element, string localName) =>
        element.GetAttributes().FirstOrDefault(a => a.NamespaceUri == W && a.LocalName == localName).Value;

    private static T Data<T>(ToolResponse response)
    {
        Assert.True(response.Ok, response.Error?.Message);
        return Assert.IsType<T>(response.Data);
    }

    // ---- set_layout_theme ----

    [Fact]
    public void Setting_slots_recolours_every_reference_to_them_and_only_those()
    {
        // The themable Blue invoice prints coral red: its theme's accent1 is F24F4F. One slot makes it blue.
        using var repo = new TargetFolder();
        repo.AppJson(".", "29.0.0.0");
        var path = repo.CopyIn(Corpus.Path(Corpus.SalesInvoiceBlueSimpleThemable), "Invoice.docx");
        var before = ThemeReferences(path);

        var result = Data<ThemeEditResultDto>(ThemeTools.SetLayoutTheme(
            path, new Dictionary<string, string> { ["accent1"] = "0070C0", ["Text2"] = "#44546a" }));

        Assert.Equal([("accent1", "F24F4F", "0070C0"), ("text2", "4C483D", "44546A")], result.Colors.Select(c => (c.Slot, c.Previous, c.Value)));

        var after = ThemeReferences(path);
        Assert.Equal(before.Count, after.Count);
        var recoloured = 0;
        for (var i = 0; i < after.Count; i++)
        {
            var (slot, tint, shade, cached) = after[i];
            if (slot is "accent1" or "text2")
            {
                Assert.Equal(ThemePalette.ApplyTintShade(slot == "accent1" ? "0070C0" : "44546A", tint, shade), cached);
                recoloured++;
            }
            else
            {
                Assert.Equal(before[i].Cached, cached);
            }
        }

        Assert.Equal(recoloured, result.RecachedReferences);
        Assert.True(recoloured > 59, $"only {recoloured} references: accent1 alone colours 26 runs, 8 fills and 25 rules");
        Assert.Contains("accent1 F24F4F -> 0070C0", result.Summary, StringComparison.Ordinal);

        using var doc = WordprocessingDocument.Open(path, false);
        Assert.Equal("0070C0", ThemePalette.Resolve(doc.MainDocumentPart!, "accent1"));
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc));
    }

    [Fact]
    public void Fonts_are_set_and_reported_with_what_they_were()
    {
        using var repo = new TargetFolder();
        repo.AppJson(".", "29.0.0.0");
        var path = repo.CopyIn(Corpus.Path(Corpus.SalesInvoiceBlueSimpleThemable), "Invoice.docx");

        var result = Data<ThemeEditResultDto>(ThemeTools.SetLayoutTheme(path, majorFont: "Segoe UI Semibold", minorFont: " Georgia "));

        Assert.Equal(("Segoe UI Light", "Segoe UI Semibold"), (result.MajorFont!.Previous, result.MajorFont.Value));
        Assert.Equal(("Segoe UI", "Georgia"), (result.MinorFont!.Previous, result.MinorFont.Value));
        Assert.Empty(result.Colors);
        Assert.Equal(0, result.RecachedReferences);

        using var doc = WordprocessingDocument.Open(path, false);
        var fonts = doc.MainDocumentPart!.ThemePart!.Theme!.ThemeElements!.FontScheme!;
        Assert.Equal("Segoe UI Semibold", fonts.MajorFont!.LatinFont!.Typeface!.Value);
        Assert.Equal("Georgia", fonts.MinorFont!.LatinFont!.Typeface!.Value);
    }

    [Fact]
    public void A_system_colour_entry_becomes_an_srgbClr_as_Microsofts_report_themes_write_it()
    {
        using var repo = new TargetFolder();
        repo.AppJson(".", "29.0.0.0");
        var path = repo.CopyIn(Corpus.Path(Corpus.StandardSalesQuote), "Quote.docx");

        Data<ThemeEditResultDto>(ThemeTools.SetLayoutTheme(path, new Dictionary<string, string> { ["text1"] = "242424" }));

        using var doc = WordprocessingDocument.Open(path, false);
        var dark1 = doc.MainDocumentPart!.ThemePart!.Theme!.ThemeElements!.ColorScheme!.Dark1Color!;
        Assert.Null(dark1.GetFirstChild<A.SystemColor>());
        Assert.Equal("242424", dark1.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
    }

    [Fact]
    public void A_slot_set_to_its_own_colour_changes_nothing()
    {
        using var repo = new TargetFolder();
        repo.AppJson(".", "29.0.0.0");
        var path = repo.CopyIn(Corpus.Path(Corpus.SalesInvoiceBlueSimpleThemable), "Invoice.docx");
        var before = ThemeReferences(path);

        var result = Data<ThemeEditResultDto>(ThemeTools.SetLayoutTheme(path, new Dictionary<string, string> { ["accent1"] = "f24f4f" }));

        Assert.Equal(0, result.RecachedReferences);
        Assert.Contains("accent1 F24F4F (unchanged)", result.Summary, StringComparison.Ordinal);
        Assert.Equal(before, ThemeReferences(path));
    }

    [Theory]
    [InlineData("28.0.0.0")]
    [InlineData(null)]
    public void The_theme_is_refused_below_BC_29_or_for_an_unknown_target_and_the_file_is_untouched(string? application)
    {
        using var repo = new TargetFolder();
        if (application is not null)
        {
            repo.AppJson(".", application);
        }

        var path = repo.CopyIn(Corpus.Path(Corpus.SalesInvoiceBlueSimpleThemable), "Invoice.docx");
        var before = File.ReadAllBytes(path);

        var response = ThemeTools.SetLayoutTheme(path, new Dictionary<string, string> { ["accent1"] = "0070C0" });

        Assert.False(response.Ok);
        Assert.Equal(ErrorCodes.BcVersionUnsupported, response.Error!.Code);
        Assert.Equal(before, File.ReadAllBytes(path));
    }

    [Theory]
    [InlineData(null, null, null, "nothing to change", "colors")]
    [InlineData("accent7", "000000", null, "not a theme colour slot", "colors")]
    [InlineData("accent1", "auto", null, "must be six hex digits RRGGBB", "colors")]
    [InlineData("accent1", "0070C", null, "must be six hex digits RRGGBB", "colors")]
    [InlineData(null, null, "   ", "must be a font family name", "majorFont")]
    public void Malformed_arguments_are_refused(string? slot, string? value, string? majorFont, string fragment, string paramName)
    {
        using var doc = WordprocessingDocument.Open(Corpus.Path(Corpus.SalesInvoiceBlueSimpleThemable), false);
        var colors = slot is null ? null : new Dictionary<string, string> { [slot] = value! };

        var ex = Assert.Throws<ArgumentException>(() => ThemeEditor.SetTheme(doc, colors, majorFont, minorFont: null));

        Assert.Contains(fragment, ex.Message, StringComparison.Ordinal);
        Assert.Equal(paramName, ex.ParamName);
    }

    [Fact]
    public void A_slot_named_twice_in_different_case_is_refused()
    {
        using var doc = WordprocessingDocument.Open(Corpus.Path(Corpus.SalesInvoiceBlueSimpleThemable), false);

        var ex = Assert.Throws<ArgumentException>(() => ThemeEditor.SetTheme(
            doc, new Dictionary<string, string> { ["accent1"] = "0070C0", ["ACCENT1"] = "002060" }, null, null));

        Assert.Contains("names 'accent1' twice", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Two_slots_mapped_to_one_theme_colour_must_agree()
    {
        var path = Corpus.UnlinkedCopy(Corpus.SalesInvoiceBlueSimpleThemable);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var mapping = doc.MainDocumentPart!.DocumentSettingsPart!.Settings!.GetFirstChild<ColorSchemeMapping>()!;
            mapping.Accent2 = ColorSchemeIndexValues.Accent1;

            var ex = Assert.Throws<ArgumentException>(() => ThemeEditor.SetTheme(
                doc, new Dictionary<string, string> { ["accent1"] = "0070C0", ["accent2"] = "002060" }, null, null));
            Assert.Contains("points both at the same theme colour", ex.Message, StringComparison.Ordinal);

            var agreed = ThemeEditor.SetTheme(doc, new Dictionary<string, string> { ["accent1"] = "0070C0", ["accent2"] = "0070C0" }, null, null);
            Assert.Equal(2, agreed.Colors.Count);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_layout_without_a_theme_part_is_refused_pointing_at_create_layout()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.PlainParagraph("x"));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);

            var ex = Assert.Throws<ArgumentException>(() => ThemeEditor.SetTheme(doc, null, "Segoe UI", null));

            Assert.Contains("no theme part", ex.Message, StringComparison.Ordinal);
            Assert.Contains("theme 'office'", ex.Message, StringComparison.Ordinal);
            Assert.Equal("majorFont", ex.ParamName);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void The_refusal_hints_name_the_slots_and_the_office_theme()
    {
        using var repo = new TargetFolder();
        repo.AppJson(".", "29.0.0.0");
        var path = repo.CopyIn(Corpus.Path(Corpus.SalesInvoiceBlueSimpleThemable), "Invoice.docx");

        var colors = ThemeTools.SetLayoutTheme(path, new Dictionary<string, string> { ["accent9"] = "000000" });
        Assert.Equal("invalid_argument", colors.Error!.Code);
        Assert.Contains("accent1 to accent6, hyperlink, followedHyperlink", colors.Error.Hint, StringComparison.Ordinal);

        var font = ThemeTools.SetLayoutTheme(path, minorFont: new string('x', 65));
        Assert.Equal("invalid_argument", font.Error!.Code);
        Assert.Contains("majorFont (headings) and minorFont (body text)", font.Error.Hint, StringComparison.Ordinal);
    }

    // ---- create_layout's theme ----

    /// <summary>A theme part's XML without its namespace declarations, which the SDK may move when it saves a part it read.</summary>
    private static XElement WithoutNamespaceDeclarations(ThemePart part)
    {
        var root = XDocument.Load(part.GetStream()).Root!;
        foreach (var element in root.DescendantsAndSelf())
        {
            element.Attributes().Where(a => a.IsNamespaceDeclaration).Remove();
        }

        return root;
    }

    [Fact]
    public void create_layout_with_the_office_theme_ships_the_stock_quotes_theme_part_and_mapping()
    {
        using var repo = new TargetFolder();
        repo.AppJson(".", "29.0.0.0");
        var path = repo.Path("New.docx");

        var created = Data<CreateResultDto>(LifecycleTools.CreateLayout(Corpus.Path(Corpus.StandardSalesQuote), path, theme: "Office"));

        Assert.Equal("office", created.AddedTheme);
        Assert.True(created.QuickValidation.Passed);
        using (var doc = WordprocessingDocument.Open(path, false))
        using (var quote = WordprocessingDocument.Open(Corpus.Path(Corpus.StandardSalesQuote), false))
        {
            var main = doc.MainDocumentPart!;
            Assert.True(XNode.DeepEquals(WithoutNamespaceDeclarations(main.ThemePart!), WithoutNamespaceDeclarations(quote.MainDocumentPart!.ThemePart!)));
            Assert.NotNull(main.DocumentSettingsPart!.Settings!.GetFirstChild<ColorSchemeMapping>());
            Assert.Equal("5B9BD5", ThemePalette.Resolve(main, "accent1"));
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc));
        }

        // The point of it: a theme colour resolves now, and the theme can be edited.
        var insert = EditTools.InsertText(path, "Themed", "documentEnd", format: new TextFormatArgument { Color = "accent1" });
        Assert.True(insert.Ok, insert.Error?.Message);
        var edit = Data<ThemeEditResultDto>(ThemeTools.SetLayoutTheme(path, new Dictionary<string, string> { ["accent1"] = "C00000" }));
        Assert.Equal(1, edit.RecachedReferences);
    }

    [Fact]
    public void create_layout_without_a_theme_ships_none_as_before()
    {
        using var repo = new TargetFolder();
        var path = repo.Path("New.docx");

        var created = Data<CreateResultDto>(LifecycleTools.CreateLayout(Corpus.Path(Corpus.StandardSalesQuote), path));

        Assert.Null(created.AddedTheme);
        using var doc = WordprocessingDocument.Open(path, false);
        Assert.Null(doc.MainDocumentPart!.ThemePart);
    }

    [Theory]
    [InlineData("dark", "theme must be 'office'")]
    [InlineData("office", "templatePath already has one ('Office Theme')")]
    public void create_layout_refuses_an_unknown_theme_and_a_second_theme_for_a_template(string theme, string fragment)
    {
        using var repo = new TargetFolder();
        var path = repo.Path("New.docx");

        var response = LifecycleTools.CreateLayout(
            Corpus.Path(Corpus.StandardSalesQuote), path, templatePath: Corpus.Path(Corpus.StandardSalesQuote), theme: theme);

        Assert.False(response.Ok);
        Assert.Equal("invalid_argument", response.Error!.Code);
        Assert.Contains(fragment, response.Error.Message, StringComparison.Ordinal);
        Assert.Contains("create_layout's theme is 'office'", response.Error.Hint, StringComparison.Ordinal);
        Assert.False(File.Exists(path));
    }
}
