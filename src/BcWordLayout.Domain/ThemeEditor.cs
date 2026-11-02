using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;

namespace BcWordLayout.Domain;

/// <summary>One theme colour <see cref="ThemeEditor.SetTheme"/> changed.</summary>
/// <param name="Slot">The slot as the caller named it, e.g. <c>accent1</c>.</param>
/// <param name="Previous">Its <c>RRGGBB</c> before, or null when the scheme had none.</param>
/// <param name="Value">Its <c>RRGGBB</c> now.</param>
public sealed record ThemeColorChange(string Slot, string? Previous, string Value);

/// <summary>One theme font <see cref="ThemeEditor.SetTheme"/> changed.</summary>
public sealed record ThemeFontChange(string? Previous, string Value);

/// <summary>What <see cref="ThemeEditor.SetTheme"/> did.</summary>
public sealed class ThemeEditResult
{
    /// <summary>Always <c>set_layout_theme</c>.</summary>
    public string Operation { get; init; } = "set_layout_theme";

    /// <summary>The slots set, in the order given.</summary>
    public required IReadOnlyList<ThemeColorChange> Colors { get; init; }

    /// <summary>The heading font (<c>a:majorFont</c>), when set.</summary>
    public ThemeFontChange? MajorFont { get; init; }

    /// <summary>The body font (<c>a:minorFont</c>), when set.</summary>
    public ThemeFontChange? MinorFont { get; init; }

    /// <summary>The theme-coloured references (text, fills, rules, underlines) whose cached colour was rewritten.</summary>
    public required int RecachedReferences { get; init; }

    /// <summary>One paragraph for the caller.</summary>
    public required string Summary { get; init; }
}

/// <summary>
/// Changes a layout's own theme (<c>word/theme/theme1.xml</c>): the colours its scheme gives the theme slots and its
/// heading and body fonts (GitHub issue #150). A layout coloured through theme slots (<c>accent1</c>, #145) prints
/// whatever its theme says: the stock themable Blue invoice prints coral red because its theme's <c>accent1</c> is
/// <c>F24F4F</c>, and making it blue is one slot.
/// </summary>
/// <remarks>
/// <para>
/// SLOTS are the ones <see cref="ThemePalette.Slots"/> lists, mapped to the scheme's entries through the document's
/// <c>w:clrSchemeMapping</c>, as <see cref="ThemePalette.Resolve"/> reads them. A new colour is written as
/// <c>a:srgbClr</c>, also where the entry held a system colour (<c>dk1</c>/<c>lt1</c> are <c>windowText</c>/
/// <c>window</c> in every corpus theme): Microsoft's own report themes Calm, Default and Playful write their
/// customised <c>dk1</c>/<c>lt1</c> that way. The scheme keeps its name.
/// </para>
/// <para>
/// THE CACHE. Word writes the colour a theme reference resolves to beside it (<c>w:color w:val</c>, <c>w:shd
/// w:fill</c>, a border's <c>w:color</c>, tint or shade applied) and recomputes it from the theme when it opens the
/// file. Every reference to a changed entry, in the content parts, the styles and the numbering, gets the new value
/// (<see cref="ThemePalette.ApplyTintShade"/>), so a renderer that reads only the cached colour prints the new one
/// too. Whether a Business Central report theme overrides a layout's own theme is open (GitHub issue #142).
/// </para>
/// </remarks>
public static class ThemeEditor
{
    private const int MaxFontNameLength = 64;

    /// <summary>
    /// Sets the slots in <paramref name="colors"/> (slot to <c>RRGGBB</c>) and the heading (<paramref name="majorFont"/>)
    /// and body (<paramref name="minorFont"/>) fonts of <paramref name="doc"/>'s theme; null leaves each alone.
    /// </summary>
    /// <exception cref="ArgumentException">
    /// Nothing to change; an unknown slot or a value that is not <c>RRGGBB</c>; two slots that name the same scheme entry
    /// with different colours; a font name that is blank, longer than 64 characters or holds control characters; or a
    /// layout with no theme part.
    /// </exception>
    public static ThemeEditResult SetTheme(
        WordprocessingDocument doc, IReadOnlyDictionary<string, string>? colors, string? majorFont, string? minorFont)
    {
        ArgumentNullException.ThrowIfNull(doc);

        var main = doc.MainDocumentPart ?? throw new InvalidDataException("Layout has no main document part.");
        if ((colors is null || colors.Count == 0) && majorFont is null && minorFont is null)
        {
            throw new ArgumentException(
                "set_layout_theme has nothing to change: give colors (theme slot to RRGGBB, e.g. {\"accent1\": "
                + "\"0070C0\"}), majorFont (the heading font) or minorFont (the body font).",
                nameof(colors));
        }

        var requested = ParseColors(colors);
        var major = majorFont is null ? null : CheckFontName(majorFont, nameof(majorFont));
        var minor = minorFont is null ? null : CheckFontName(minorFont, nameof(minorFont));

        var theme = main.ThemePart?.Theme;
        var scheme = theme?.ThemeElements?.ColorScheme;
        var fonts = theme?.ThemeElements?.FontScheme;
        if (theme is null || scheme is null || fonts is null)
        {
            throw new ArgumentException(
                "This layout has no theme part (word/theme/theme1.xml), so it has no theme colours or fonts to set. A "
                + "blank create_layout build gets the stock Office theme with theme 'office'; a layout built on a "
                + "template keeps the template's theme.",
                requested.Count > 0 ? nameof(colors) : major is not null ? nameof(majorFont) : nameof(minorFont));
        }

        // Two slots may name one scheme entry (a custom w:clrSchemeMapping); they must then agree.
        var byEntry = new Dictionary<ColorSchemeIndexValues, (string Slot, string Rgb)>();
        foreach (var (slot, rgb) in requested)
        {
            var index = ThemePalette.SchemeIndex(main, slot)!.Value;
            if (byEntry.TryGetValue(index, out var other) && other.Rgb != rgb)
            {
                throw new ArgumentException(
                    $"colors sets '{other.Slot}' to {other.Rgb} and '{slot}' to {rgb}, but this layout's colour mapping "
                    + $"(w:clrSchemeMapping) points both at the same theme colour ('{index}'). Give them one value.",
                    nameof(colors));
            }

            byEntry[index] = (slot, rgb);
        }

        var changes = new List<ThemeColorChange>();
        var changed = new Dictionary<ColorSchemeIndexValues, string>();
        foreach (var (slot, rgb) in requested)
        {
            var index = ThemePalette.SchemeIndex(main, slot)!.Value;
            var entry = ThemePalette.Entry(scheme, index);
            var previous = ThemePalette.EntryColor(entry);
            changes.Add(new ThemeColorChange(slot, previous, rgb));
            if (entry is null || previous == rgb)
            {
                continue;
            }

            entry.RemoveAllChildren();
            entry.AppendChild(new A.RgbColorModelHex { Val = rgb });
            changed[index] = rgb;
        }

        var majorChange = major is null ? null : SetFont(fonts.MajorFont, major);
        var minorChange = minor is null ? null : SetFont(fonts.MinorFont, minor);
        theme.Save();

        var recached = changed.Count == 0 ? 0 : Recache(main, changed);
        return new ThemeEditResult
        {
            Colors = changes,
            MajorFont = majorChange,
            MinorFont = minorChange,
            RecachedReferences = recached,
            Summary = Summarize(theme, changes, changed.Count, majorChange, minorChange, recached, main),
        };
    }

    private static List<(string Slot, string Rgb)> ParseColors(IReadOnlyDictionary<string, string>? colors)
    {
        var parsed = new List<(string Slot, string Rgb)>();
        foreach (var (key, value) in colors ?? new Dictionary<string, string>())
        {
            var slot = ThemePalette.Slots.FirstOrDefault(s => string.Equals(s, key?.Trim(), StringComparison.OrdinalIgnoreCase))
                ?? throw new ArgumentException(
                    $"colors has '{key}', which is not a theme colour slot; use {string.Join(", ", ThemePalette.Slots)}.",
                    nameof(colors));
            var hex = value?.Trim() is { } trimmed && trimmed.StartsWith('#') ? trimmed[1..] : value?.Trim();
            if (hex is not { Length: 6 } || !hex.All(Uri.IsHexDigit))
            {
                throw new ArgumentException(
                    $"colors.{slot} must be six hex digits RRGGBB (e.g. '0070C0'; a leading '#' is fine); got '{value}'. "
                    + "A theme colour is a colour, not 'auto' or another theme colour.",
                    nameof(colors));
            }

            if (parsed.Any(p => p.Slot == slot))
            {
                throw new ArgumentException($"colors names '{slot}' twice.", nameof(colors));
            }

            parsed.Add((slot, hex.ToUpperInvariant()));
        }

        return parsed;
    }

    private static string CheckFontName(string fontName, string paramName)
    {
        var trimmed = fontName.Trim();
        if (trimmed.Length == 0 || trimmed.Length > MaxFontNameLength || trimmed.Any(char.IsControl))
        {
            throw new ArgumentException(
                $"{paramName} must be a font family name such as 'Segoe UI' (1 to {MaxFontNameLength} characters, no "
                + $"control characters); got '{fontName}'.",
                paramName);
        }

        return trimmed;
    }

    private static ThemeFontChange SetFont(A.FontCollectionType? font, string name)
    {
        var latin = font?.LatinFont;
        var previous = latin?.Typeface?.Value;
        if (font is not null)
        {
            if (latin is null)
            {
                latin = new A.LatinFont();
                font.PrependChild(latin);
            }

            latin.Typeface = name;
        }

        return new ThemeFontChange(string.IsNullOrEmpty(previous) ? null : previous, name);
    }

    /// <summary>
    /// Rewrites the colour cached beside every theme reference to a <paramref name="changed"/> entry, in the content
    /// parts, the styles and the numbering. Returns how many references it rewrote.
    /// </summary>
    private static int Recache(MainDocumentPart main, Dictionary<ColorSchemeIndexValues, string> changed)
    {
        var roots = PartWalker.ContentParts(main).Select(p => p.Root).ToList();
        if (main.StyleDefinitionsPart?.Styles is { } styles)
        {
            roots.Add(styles);
        }

        if (main.NumberingDefinitionsPart?.Numbering is { } numbering)
        {
            roots.Add(numbering);
        }

        var count = 0;
        foreach (var root in roots)
        {
            foreach (var element in root.Descendants())
            {
                switch (element)
                {
                    case Color color when NewColor(color.ThemeColor, color.ThemeTint, color.ThemeShade) is { } rgb:
                        color.Val = rgb;
                        count++;
                        break;
                    case Shading shading:
                        if (NewColor(shading.ThemeFill, shading.ThemeFillTint, shading.ThemeFillShade) is { } fill)
                        {
                            shading.Fill = fill;
                            count++;
                        }

                        if (NewColor(shading.ThemeColor, shading.ThemeTint, shading.ThemeShade) is { } pattern)
                        {
                            shading.Color = pattern;
                            count++;
                        }

                        break;
                    case BorderType border when NewColor(border.ThemeColor, border.ThemeTint, border.ThemeShade) is { } rule:
                        border.Color = rule;
                        count++;
                        break;
                    case Underline underline when NewColor(underline.ThemeColor, underline.ThemeTint, underline.ThemeShade) is { } line:
                        underline.Color = line;
                        count++;
                        break;
                }
            }
        }

        foreach (var part in main.HeaderParts.Select(h => (OpenXmlPartRootElement?)h.Header)
                     .Concat(main.FooterParts.Select(f => (OpenXmlPartRootElement?)f.Footer))
                     .Append(main.Document)
                     .Append(main.StyleDefinitionsPart?.Styles)
                     .Append(main.NumberingDefinitionsPart?.Numbering)
                     .OfType<OpenXmlPartRootElement>())
        {
            part.Save();
        }

        return count;

        string? NewColor(OpenXmlSimpleType? themeColor, StringValue? tint, StringValue? shade) =>
            ThemePalette.SchemeIndex(main, themeColor?.InnerText) is { } index && changed.TryGetValue(index, out var rgb)
                ? ThemePalette.ApplyTintShade(rgb, tint?.Value, shade?.Value)
                : null;
    }

    private static string Summarize(
        A.Theme theme,
        List<ThemeColorChange> colors,
        int changedEntries,
        ThemeFontChange? major,
        ThemeFontChange? minor,
        int recached,
        MainDocumentPart main)
    {
        var name = theme.Name?.Value is { Length: > 0 } themeName ? $"'{themeName}'" : "(unnamed)";
        var parts = new List<string>();
        if (colors.Count > 0)
        {
            parts.Add("colours " + string.Join(", ", colors.Select(c =>
                c.Previous == c.Value ? $"{c.Slot} {c.Value} (unchanged)" : $"{c.Slot} {c.Previous ?? "(none)"} -> {c.Value}")));
        }

        if (major is not null)
        {
            parts.Add($"heading font {Quote(major.Previous)} -> '{major.Value}'");
        }

        if (minor is not null)
        {
            parts.Add($"body font {Quote(minor.Previous)} -> '{minor.Value}'");
        }

        var summary = $"Set this layout's theme {name}: {string.Join("; ", parts)}.";
        if (changedEntries > 0)
        {
            summary += recached == 0
                ? " Nothing in the layout uses the changed colours yet: text, fills and rules given them as a theme colour "
                  + "(format.color, fill, borderColor 'accent1') follow the theme."
                : $" The {recached.ToString(CultureInfo.InvariantCulture)} text colour(s), fill(s) and rule(s) that use "
                  + "them follow the theme; the colour each caches beside its theme reference was updated too, so a "
                  + "renderer that reads only that prints the new colour.";
        }

        return summary
            + BcOnlineFonts.SummaryNote(major?.Value, main)
            + BcOnlineFonts.SummaryNote(minor?.Value, main);

        static string Quote(string? font) => font is null ? "(none)" : $"'{font}'";
    }
}
