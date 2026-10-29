using System.Globalization;
using BcWordLayout.Domain.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Domain;

/// <summary>
/// The one place a <see cref="TextFormat"/> is checked and written (GitHub issue #15). Every tool that formats
/// text goes through it: the insert tools, <c>set_cell_text</c>, the row tools and <c>set_text_format</c>, so a
/// member means the same thing on each of them.
/// </summary>
/// <remarks>
/// <para>
/// WHAT EACH MEMBER WRITES, and the corpus evidence for it (ADR-0005; counts over the ten layouts the corpus held
/// when GitHub issue #15 was designed, direct runs plus the controls' own <c>w:sdtPr/w:rPr</c>):
/// <list type="bullet">
/// <item><c>Bold</c>: <c>w:b</c>. False removes it, or writes <c>w:b w:val="0"</c> when the run carries a
/// character style, since removing <c>w:b</c> cannot un-bold <c>Strong</c> (GitHub issue #122). The explicit off
/// is in the corpus 55 times.</item>
/// <item><c>Italic</c>: <c>w:i</c> + <c>w:iCs</c>; false as for bold. The corpus applies italic only through
/// style definitions (<c>Emphasis</c>, <c>SubtleEmphasis</c>, <c>IntenseEmphasis</c>, <c>Heading5Char</c>/
/// <c>7Char</c>/<c>9Char</c> in StandardSalesInvoiceVatSpec), always as that pair; sandbox probe P13 checks
/// the direct form in Business Central.</item>
/// <item><c>Underline</c>: <c>w:u w:val="single"</c> (26 times in 3 layouts); false removes it, or writes
/// <c>w:val="none"</c> (twice in the corpus) over a character style.</item>
/// <item><c>FontSizePoints</c>: <c>w:sz</c> + <c>w:szCs</c> in half-points, as Word writes them.</item>
/// <item><c>FontName</c>: <c>w:rFonts</c> with <c>w:ascii</c>/<c>w:hAnsi</c>/<c>w:cs</c> all naming the font,
/// the shape of ~700 entries in 8 layouts (nearly all Segoe UI). Theme-font attributes are dropped, because
/// Word lets them win over the names; the East Asian font and the hint are kept. Any name is written, but the
/// tool summary warns when Business Central online lacks the font (<see cref="FontNote"/>).</item>
/// <item><c>Color</c>: <c>w:color w:val</c> as <c>RRGGBB</c> or <c>auto</c> (<c>0070C0</c> 23 times,
/// <c>auto</c> 7, in 3 layouts). Theme-colour attributes are dropped for the same reason as theme fonts;
/// writing theme colours is out of scope (two witnesses, and they depend on the theme part).</item>
/// <item><c>CharacterStyle</c>: <c>w:rStyle</c> (<c>Strong</c> 20 times, <c>PageNumber</c>,
/// <c>Heading1Char</c> and others, in 4 layouts), refused unless the layout's styles part defines it as a
/// CHARACTER style: a dangling or paragraph-style reference does nothing in any renderer.</item>
/// </list>
/// </para>
/// <para>
/// Every element is placed with <see cref="OpenXmlCompositeElement.AddChild"/>, which puts it at its
/// schema-ordered position, so the same code serves a run's <c>w:rPr</c>, a control's <c>w:sdtPr/w:rPr</c> and a
/// paragraph mark's <c>w:pPr/w:rPr</c> whatever they already contain.
/// </para>
/// </remarks>
internal static class RunFormatter
{
    /// <summary>A <see cref="TextFormat"/> that has passed <see cref="Prepare"/>: values normalized for writing.</summary>
    internal sealed record Prepared(
        bool? Bold,
        bool? Italic,
        bool? Underline,
        string? HalfPoints,
        string? FontName,
        string? Color,
        string? CharacterStyle);

    /// <summary>The longest font name accepted; no installed font family comes close.</summary>
    private const int MaxFontNameLength = 64;

    /// <summary>
    /// Checks <paramref name="format"/> against <paramref name="doc"/> and normalizes it for writing, or returns
    /// null when it is null or sets nothing. Called before any mutation, so a bad member leaves the layout
    /// untouched.
    /// </summary>
    /// <exception cref="ArgumentException">
    /// A member is out of range or malformed, or <see cref="TextFormat.CharacterStyle"/> names no character style
    /// the layout defines (ParamName <c>format</c>).
    /// </exception>
    internal static Prepared? Prepare(TextFormat? format, WordprocessingDocument doc)
    {
        ArgumentNullException.ThrowIfNull(doc);

        if (format is null || format.IsEmpty)
        {
            return null;
        }

        return new Prepared(
            format.Bold,
            format.Italic,
            format.Underline,
            format.FontSizePoints is { } points ? HalfPoints(points, nameof(format)) : null,
            format.FontName is { } font ? CheckFontName(font, nameof(format)) : null,
            format.Color is { } color ? NormalizeColor(color, "format.color", nameof(format)) : null,
            format.CharacterStyle is { } style ? CheckCharacterStyle(style, doc, nameof(format)) : null);
    }

    /// <summary>
    /// Normalizes a colour argument to what OOXML stores: <c>auto</c>, or six upper-case hex digits (a leading
    /// <c>#</c> is accepted). Shared with <c>set_cell_format</c>'s fill, which takes the same values.
    /// </summary>
    /// <exception cref="ArgumentException">Anything else, naming <paramref name="what"/> and keyed on <paramref name="paramName"/>.</exception>
    internal static string NormalizeColor(string value, string what, string paramName)
    {
        var trimmed = value.Trim();
        if (string.Equals(trimmed, "auto", StringComparison.OrdinalIgnoreCase))
        {
            return "auto";
        }

        var hex = trimmed.StartsWith('#') ? trimmed[1..] : trimmed;
        if (hex.Length == 6 && hex.All(Uri.IsHexDigit))
        {
            return hex.ToUpperInvariant();
        }

        throw new ArgumentException(
            $"{what} must be a colour as six hex digits RRGGBB (e.g. '0070C0'; a leading '#' is fine) or 'auto' "
            + $"(got '{value}').",
            paramName);
    }

    /// <summary>
    /// Writes <paramref name="format"/> onto <paramref name="properties"/>: a run's <c>w:rPr</c>, a control's
    /// <c>w:sdtPr/w:rPr</c>, or a paragraph mark's <c>w:pPr/w:rPr</c>. Members left null are not touched.
    /// </summary>
    internal static void Apply(OpenXmlCompositeElement properties, Prepared format)
    {
        ArgumentNullException.ThrowIfNull(properties);
        ArgumentNullException.ThrowIfNull(format);

        // The style first: the explicit-off rule below looks at whether a style is present.
        if (format.CharacterStyle is { } style)
        {
            Set(properties, new RunStyle { Val = style });
        }

        var styled = properties.GetFirstChild<RunStyle>() is not null;

        if (format.Bold is { } bold)
        {
            Set(properties, bold ? new Bold() : styled ? new Bold { Val = false } : null);
        }

        if (format.Italic is { } italic)
        {
            Set(properties, italic ? new Italic() : styled ? new Italic { Val = false } : null);
            Set(properties, italic ? new ItalicComplexScript() : null);
        }

        if (format.Underline is { } underline)
        {
            Set(properties, underline
                ? new Underline { Val = UnderlineValues.Single }
                : styled ? new Underline { Val = UnderlineValues.None } : null);
        }

        if (format.HalfPoints is { } halfPoints)
        {
            Set(properties, new FontSize { Val = halfPoints });
            Set(properties, new FontSizeComplexScript { Val = halfPoints });
        }

        if (format.FontName is { } font)
        {
            var existing = properties.GetFirstChild<RunFonts>();
            var fonts = new RunFonts { Ascii = font, HighAnsi = font, ComplexScript = font };
            if (existing?.EastAsia is { } eastAsia)
            {
                fonts.EastAsia = (StringValue)eastAsia.Clone();
            }

            if (existing?.EastAsiaTheme is { } eastAsiaTheme)
            {
                fonts.EastAsiaTheme = (EnumValue<ThemeFontValues>)eastAsiaTheme.Clone();
            }

            if (existing?.Hint is { } hint)
            {
                fonts.Hint = (EnumValue<FontTypeHintValues>)hint.Clone();
            }

            Set(properties, fonts);
        }

        if (format.Color is { } color)
        {
            Set(properties, new Color { Val = color });
        }
    }

    /// <summary>
    /// The sentence a tool summary carries when <paramref name="format"/> names a font Business Central online
    /// will not print as named (see <see cref="BcOnlineFonts.SummaryNote"/>); empty otherwise.
    /// </summary>
    internal static string FontNote(Prepared? format, WordprocessingDocument doc) =>
        BcOnlineFonts.SummaryNote(format?.FontName, doc.MainDocumentPart);

    /// <summary>
    /// The members <paramref name="format"/> sets, as a short phrase for a tool summary ("bold, not italic,
    /// 8 pt, Segoe UI, colour 0070C0, style 'Strong'").
    /// </summary>
    internal static string Describe(Prepared format)
    {
        var parts = new List<string>();
        if (format.CharacterStyle is { } style)
        {
            parts.Add($"character style '{style}'");
        }

        if (format.Bold is { } bold)
        {
            parts.Add(bold ? "bold" : "not bold");
        }

        if (format.Italic is { } italic)
        {
            parts.Add(italic ? "italic" : "not italic");
        }

        if (format.Underline is { } underline)
        {
            parts.Add(underline ? "underlined" : "not underlined");
        }

        if (format.HalfPoints is { } halfPoints)
        {
            var points = int.Parse(halfPoints, CultureInfo.InvariantCulture) / 2.0;
            parts.Add($"{points.ToString(CultureInfo.InvariantCulture)} pt");
        }

        if (format.FontName is { } font)
        {
            parts.Add(font);
        }

        if (format.Color is { } color)
        {
            parts.Add($"colour {color}");
        }

        return string.Join(", ", parts);
    }

    /// <summary>
    /// Replaces every child of type <typeparamref name="T"/> with <paramref name="value"/> (or just removes them
    /// when it is null), placing the new child at its schema-ordered position.
    /// </summary>
    private static void Set<T>(OpenXmlCompositeElement properties, T? value)
        where T : OpenXmlElement
    {
        properties.RemoveAllChildren<T>();
        if (value is not null)
        {
            properties.AddChild(value, throwOnError: true);
        }
    }

    private static string HalfPoints(double points, string paramName)
    {
        if (double.IsNaN(points) || points is < 4 or > 96)
        {
            throw new ArgumentException(
                $"format.fontSizePoints must be between 4 and 96 (got {points.ToString(CultureInfo.InvariantCulture)}).",
                paramName);
        }

        // w:sz/w:szCs are half-points; Word itself writes both.
        return ((int)Math.Round(points * 2)).ToString(CultureInfo.InvariantCulture);
    }

    private static string CheckFontName(string fontName, string paramName)
    {
        var trimmed = fontName.Trim();
        if (trimmed.Length == 0 || trimmed.Length > MaxFontNameLength || trimmed.Any(char.IsControl))
        {
            throw new ArgumentException(
                $"format.fontName must be a font family name such as 'Segoe UI' (1 to {MaxFontNameLength} "
                + $"characters, no control characters); got '{fontName}'.",
                paramName);
        }

        return trimmed;
    }

    /// <summary>
    /// Refuses a style id the layout does not define as a CHARACTER style (ids are case-sensitive, as in Word's
    /// own model), naming the ones it does.
    /// </summary>
    private static string CheckCharacterStyle(string styleId, WordprocessingDocument doc, string paramName)
    {
        var styles = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles?.Elements<Style>().ToList() ?? [];
        var match = styles.FirstOrDefault(s => string.Equals(s.StyleId?.Value, styleId, StringComparison.Ordinal));
        if (match?.Type?.Value == StyleValues.Character)
        {
            return styleId;
        }

        var characterStyles = styles
            .Where(s => s.Type?.Value == StyleValues.Character && s.StyleId?.Value is { Length: > 0 })
            .Select(s => s.StyleId!.Value!)
            .Where(id => id != "DefaultParagraphFont")
            .ToList();
        var available = characterStyles.Count == 0
            ? "This layout defines no character styles: start from a templatePath that defines the style, or "
              + "set bold/italic/fontName/color directly."
            : $"Character styles this layout defines: {string.Join(", ", characterStyles.Select(s => $"'{s}'"))}.";
        var reason = match is null
            ? "the layout's styles part does not define it"
            : $"it is a {match.Type?.InnerText ?? "typeless"} style, not a character style";

        throw new ArgumentException(
            $"format.characterStyle '{styleId}' cannot be applied: {reason}, so the reference would do nothing. "
            + available,
            paramName);
    }
}
