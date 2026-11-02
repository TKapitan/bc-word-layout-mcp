using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;

namespace BcWordLayout.Domain;

/// <summary>
/// A layout's embedded theme as the tools see it (GitHub issues #145, #149): its colour scheme read through the
/// document's <c>w:clrSchemeMapping</c>, keyed by the slot names <c>w:themeColor</c> and <c>w:themeFill</c> use, and its
/// major and minor Latin fonts.
/// </summary>
/// <param name="Name">The theme's own name (<c>a:theme/@name</c>), e.g. <c>Office Theme</c>.</param>
/// <param name="ColorScheme">The colour scheme's name, e.g. <c>Business Invoice</c>.</param>
/// <param name="Colors">Each slot of <see cref="ThemePalette.Slots"/> the scheme defines, as <c>RRGGBB</c>, in that order.</param>
/// <param name="MajorFont">The heading font (<c>a:majorFont/a:latin</c>), or null when the theme names none.</param>
/// <param name="MinorFont">The body font (<c>a:minorFont/a:latin</c>), or null when the theme names none.</param>
public sealed record ThemeInfo(
    string? Name, string? ColorScheme, IReadOnlyList<KeyValuePair<string, string>> Colors, string? MajorFont, string? MinorFont);

/// <summary>
/// Reads the colour scheme of a layout's theme part (<c>word/theme/theme1.xml</c>) and resolves a theme slot to the
/// colour Word shows for it, tint or shade included.
/// </summary>
/// <remarks>
/// <para>
/// SLOTS. A document names a theme colour by the slots of <c>ST_ThemeColor</c> that Word's palette offers: <c>text1</c>,
/// <c>background1</c>, <c>text2</c>, <c>background2</c>, <c>accent1</c> to <c>accent6</c>, <c>hyperlink</c> and
/// <c>followedHyperlink</c>. The settings part's <c>w:clrSchemeMapping</c> maps each to an entry of the theme's
/// <c>a:clrScheme</c> (<c>t1</c> to <c>dark1</c>, <c>bg1</c> to <c>light1</c> and so on); every corpus layout carries
/// the default mapping, which applies when the element is absent.
/// </para>
/// <para>
/// TINTS AND SHADES change the colour's HSL luminance (ECMA-376 Part 1, 17.3.2.6): a tint <c>t</c> (the hex byte over
/// 255) gives <c>L * t + (1 - t)</c>, a shade gives <c>L * t</c>. Word caches the result in <c>w:val</c> /
/// <c>w:fill</c> / <c>w:color</c> beside the theme attributes. Computed in floating point with each channel truncated
/// to a byte, this gives 14 of the 18 distinct tinted or shaded values the corpus caches exactly, the rule colours
/// included (text2 lighter 40% is <c>9B9482</c>, background2 darker 25% <c>AEAAAA</c>), and the other four one step
/// off in one channel; rounding instead gives only 9. Word's own arithmetic is not documented. Word recomputes the
/// cache from the theme when it opens the file, so the step only matters to a renderer that ignores the theme
/// attributes.
/// </para>
/// </remarks>
public static class ThemePalette
{
    /// <summary>The slots a colour argument may name, as <c>w:themeColor</c> spells them, in Word's palette order.</summary>
    public static IReadOnlyList<string> Slots { get; } =
    [
        "text1", "background1", "text2", "background2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
        "hyperlink", "followedHyperlink",
    ];

    /// <summary>
    /// The theme of <paramref name="main"/>, or null when the layout has no theme part (a blank <c>create_layout</c>
    /// build: <see cref="DefaultStylesScaffold"/> ships styles only).
    /// </summary>
    public static ThemeInfo? Read(MainDocumentPart main)
    {
        ArgumentNullException.ThrowIfNull(main);

        if (main.ThemePart?.Theme is not { } theme)
        {
            return null;
        }

        var scheme = theme.ThemeElements?.ColorScheme;
        var colors = new List<KeyValuePair<string, string>>();
        foreach (var slot in Slots)
        {
            if (Resolve(main, slot) is { } rgb)
            {
                colors.Add(new(slot, rgb));
            }
        }

        var fonts = theme.ThemeElements?.FontScheme;
        return new ThemeInfo(
            theme.Name?.Value,
            scheme?.Name?.Value,
            colors,
            NullIfEmpty(fonts?.MajorFont?.LatinFont?.Typeface?.Value),
            NullIfEmpty(fonts?.MinorFont?.LatinFont?.Typeface?.Value));
    }

    /// <summary>
    /// The <c>RRGGBB</c> colour <paramref name="slot"/> (one of <see cref="Slots"/>) has in <paramref name="main"/>'s
    /// theme, through the document's colour-scheme mapping; null when the layout has no theme part or the scheme lacks
    /// the entry.
    /// </summary>
    public static string? Resolve(MainDocumentPart main, string slot)
    {
        ArgumentNullException.ThrowIfNull(main);

        if (main.ThemePart?.Theme?.ThemeElements?.ColorScheme is not { } scheme)
        {
            return null;
        }

        var mapping = main.DocumentSettingsPart?.Settings?.GetFirstChild<ColorSchemeMapping>();
        var index = slot switch
        {
            "text1" => mapping?.Text1?.Value ?? ColorSchemeIndexValues.Dark1,
            "background1" => mapping?.Background1?.Value ?? ColorSchemeIndexValues.Light1,
            "text2" => mapping?.Text2?.Value ?? ColorSchemeIndexValues.Dark2,
            "background2" => mapping?.Background2?.Value ?? ColorSchemeIndexValues.Light2,
            "accent1" => mapping?.Accent1?.Value ?? ColorSchemeIndexValues.Accent1,
            "accent2" => mapping?.Accent2?.Value ?? ColorSchemeIndexValues.Accent2,
            "accent3" => mapping?.Accent3?.Value ?? ColorSchemeIndexValues.Accent3,
            "accent4" => mapping?.Accent4?.Value ?? ColorSchemeIndexValues.Accent4,
            "accent5" => mapping?.Accent5?.Value ?? ColorSchemeIndexValues.Accent5,
            "accent6" => mapping?.Accent6?.Value ?? ColorSchemeIndexValues.Accent6,
            "hyperlink" => mapping?.Hyperlink?.Value ?? ColorSchemeIndexValues.Hyperlink,
            "followedHyperlink" => mapping?.FollowedHyperlink?.Value ?? ColorSchemeIndexValues.FollowedHyperlink,
            _ => throw new ArgumentOutOfRangeException(nameof(slot), slot, "Not a theme colour slot."),
        };

        OpenXmlCompositeElement? entry =
            index == ColorSchemeIndexValues.Dark1 ? scheme.Dark1Color
            : index == ColorSchemeIndexValues.Light1 ? scheme.Light1Color
            : index == ColorSchemeIndexValues.Dark2 ? scheme.Dark2Color
            : index == ColorSchemeIndexValues.Light2 ? scheme.Light2Color
            : index == ColorSchemeIndexValues.Accent1 ? scheme.Accent1Color
            : index == ColorSchemeIndexValues.Accent2 ? scheme.Accent2Color
            : index == ColorSchemeIndexValues.Accent3 ? scheme.Accent3Color
            : index == ColorSchemeIndexValues.Accent4 ? scheme.Accent4Color
            : index == ColorSchemeIndexValues.Accent5 ? scheme.Accent5Color
            : index == ColorSchemeIndexValues.Accent6 ? scheme.Accent6Color
            : index == ColorSchemeIndexValues.Hyperlink ? scheme.Hyperlink
            : scheme.FollowedHyperlinkColor;

        // A system colour (dk1/lt1 are windowText/window in every corpus theme) carries the RGB it last had.
        var rgb = entry?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value
                  ?? entry?.GetFirstChild<A.SystemColor>()?.LastColor?.Value;
        return rgb is { Length: 6 } && rgb.All(Uri.IsHexDigit) ? rgb.ToUpperInvariant() : null;
    }

    /// <summary>
    /// <paramref name="rgb"/> with a theme tint or shade applied (each a hex byte, <c>99</c>, or null): the colour Word
    /// caches beside <c>w:themeTint</c> / <c>w:themeShade</c>. See the type remarks for the formula.
    /// </summary>
    public static string ApplyTintShade(string rgb, string? tint, string? shade)
    {
        ArgumentNullException.ThrowIfNull(rgb);

        if (tint is null && shade is null)
        {
            return rgb;
        }

        var r = Channel(rgb, 0);
        var g = Channel(rgb, 2);
        var b = Channel(rgb, 4);
        var (h, s, l) = ToHsl(r, g, b);
        if (tint is not null)
        {
            var t = Fraction(tint);
            l = l * t + (1 - t);
        }

        if (shade is not null)
        {
            l *= Fraction(shade);
        }

        var (r2, g2, b2) = FromHsl(h, s, l);
        return string.Create(CultureInfo.InvariantCulture, $"{Byte(r2):X2}{Byte(g2):X2}{Byte(b2):X2}");
    }

    private static double Channel(string rgb, int at) =>
        int.Parse(rgb.AsSpan(at, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture) / 255.0;

    private static double Fraction(string hexByte) =>
        int.Parse(hexByte, NumberStyles.HexNumber, CultureInfo.InvariantCulture) / 255.0;

    // Truncated, as the corpus's cached values are (see the type remarks); the epsilon keeps an exact byte such as
    // 1 - 0xD9/255 = 38/255 from landing a hair below 38 in floating point and truncating to 37.
    private static int Byte(double value) =>
        (int)Math.Clamp(Math.Floor((value * 255) + 1e-9), 0, 255);

    private static (double H, double S, double L) ToHsl(double r, double g, double b)
    {
        var max = Math.Max(r, Math.Max(g, b));
        var min = Math.Min(r, Math.Min(g, b));
        var l = (max + min) / 2;
        if (max == min)
        {
            return (0, 0, l);
        }

        var d = max - min;
        var s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
        double h;
        if (max == r)
        {
            h = ((g - b) / d) + (g < b ? 6 : 0);
        }
        else if (max == g)
        {
            h = ((b - r) / d) + 2;
        }
        else
        {
            h = ((r - g) / d) + 4;
        }

        return (h / 6, s, l);
    }

    private static (double R, double G, double B) FromHsl(double h, double s, double l)
    {
        if (s == 0)
        {
            return (l, l, l);
        }

        var q = l < 0.5 ? l * (1 + s) : l + s - (l * s);
        var p = (2 * l) - q;
        return (Hue(p, q, h + (1.0 / 3)), Hue(p, q, h), Hue(p, q, h - (1.0 / 3)));
    }

    private static double Hue(double p, double q, double t)
    {
        t -= Math.Floor(t);
        return t switch
        {
            < 1.0 / 6 => p + ((q - p) * 6 * t),
            < 1.0 / 2 => q,
            < 2.0 / 3 => p + ((q - p) * ((2.0 / 3) - t) * 6),
            _ => p,
        };
    }

    private static string? NullIfEmpty(string? value) => string.IsNullOrWhiteSpace(value) ? null : value;
}
