using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Domain;

/// <summary>What a colour argument colours: each writes different attributes, and the corpus evidence differs.</summary>
internal enum ColorUse
{
    /// <summary><c>format.color</c>: <c>w:color</c> on a run, a control or a paragraph mark.</summary>
    Text,

    /// <summary><c>set_cell_format</c>'s <c>fill</c>: <c>w:shd</c> on a cell.</summary>
    Fill,

    /// <summary><c>borderColor</c>: the colour of a cell or table rule.</summary>
    Rule,
}

/// <summary>
/// One checked colour argument (GitHub issues #143, #145): <c>auto</c>, a literal <c>RRGGBB</c>, or a theme slot such as
/// <c>accent1</c>, optionally <c>lighter N%</c> or <c>darker N%</c> as Word's palette names its tints and shades. Every
/// tool argument that takes a colour goes through <see cref="Parse"/>, so they all accept the same values.
/// </summary>
/// <remarks>
/// <para>
/// WHAT IS WRITTEN, each shape as the corpus carries it (ADR-0005):
/// <list type="bullet">
/// <item>Text: <c>&lt;w:color w:val="F24F4F" w:themeColor="accent1"/&gt;</c> (26 times in
/// StandardSalesInvoiceBlueSimpleThemable), with <c>w:themeTint</c> / <c>w:themeShade</c> for a tint or shade (text1
/// tints <c>D9</c>, <c>BF</c>, <c>A6</c>, <c>80</c> in its styles part; accent5 shade <c>BF</c> in
/// StandardSalesShipmentBlueThemable).</item>
/// <item>Fill: <c>&lt;w:shd w:val="clear" w:color="auto" w:fill="F24F4F" w:themeFill="accent1"/&gt;</c> (8 cells in the
/// same capture), with <c>w:themeFillShade</c> for a shade (background1 <c>F2</c> in four styles parts). No captured
/// layout carries <c>w:themeFillTint</c>, so a LIGHTER fill is refused until one does.</item>
/// <item>Rule: <c>w:color="F24F4F" w:themeColor="accent1"</c> on the edge (25 rules in the same capture), with
/// <c>w:themeTint</c> (text2 <c>99</c>, StandardSalesInvoiceVatSpec) or <c>w:themeShade</c> (background2 <c>BF</c>,
/// SalesInvoiceForSubscriptionBilling).</item>
/// </list>
/// The literal value beside a theme attribute is the slot's colour in the layout's own theme, tint or shade applied
/// (<see cref="ThemePalette"/>), as Word writes it. A theme slot is refused when the layout has no theme part, as
/// <c>characterStyle</c> refuses an undefined style. A literal colour drops any theme attributes the element had: Word
/// lets them win over the literal value.
/// </para>
/// <para>
/// PERCENTAGES. Word's palette shows a tint as "Lighter 40%" and stores <c>themeTint</c> = 255 × (1 − 0.40) = <c>99</c>;
/// a shade as "Darker 25%", <c>themeShade</c> = 255 × (1 − 0.25) = <c>BF</c>. Every tint and shade byte in the corpus
/// is one of those percentages (<c>33</c> lighter 80%, <c>99</c> lighter 40%, <c>80</c>/<c>7F</c> 50%, <c>A6</c> 35%,
/// <c>BF</c> 25%, <c>D9</c> 15%, <c>F2</c> darker 5%), so the argument takes the percentage as Word shows it.
/// </para>
/// </remarks>
public sealed record LayoutColor
{
    private static readonly Regex Shape = new(
        @"^\s*(?<slot>[A-Za-z]+[1-6]?)(?:\s+(?<direction>lighter|darker)\s+(?<percent>\d{1,2})\s*%?)?\s*$",
        RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    private LayoutColor(string value, string? themeSlot, string? tint, string? shade, int? percent)
    {
        Value = value;
        ThemeSlot = themeSlot;
        Tint = tint;
        Shade = shade;
        Percent = percent;
    }

    /// <summary>The literal colour written to <c>w:val</c> / <c>w:fill</c> / <c>w:color</c>: <c>auto</c> or <c>RRGGBB</c>.</summary>
    internal string Value { get; }

    /// <summary>The theme slot, as <c>w:themeColor</c> spells it, or null for a literal colour.</summary>
    internal string? ThemeSlot { get; }

    /// <summary>The tint as a hex byte (<c>99</c>), or null.</summary>
    internal string? Tint { get; }

    /// <summary>The shade as a hex byte (<c>BF</c>), or null.</summary>
    internal string? Shade { get; }

    private int? Percent { get; }

    /// <summary>True for a theme slot.</summary>
    internal bool IsTheme => ThemeSlot is not null;

    /// <summary>
    /// True when <paramref name="argument"/> names a theme slot, with or without a tint or shade, whether or not the
    /// rest of it is well formed: what the tools check before <see cref="BcFeatures.ThemableFormatting"/>'s gate.
    /// </summary>
    public static bool NamesThemeSlot(string? argument) =>
        argument is not null
        && Shape.Match(argument) is { Success: true } match
        && SlotName(match.Groups["slot"].Value) is not null;

    /// <summary>
    /// Checks <paramref name="argument"/> and resolves it against <paramref name="doc"/>'s theme for <paramref name="use"/>.
    /// </summary>
    /// <param name="what">The argument's name in messages, e.g. <c>format.color</c>.</param>
    /// <param name="paramName">The parameter the refusal is keyed on, for its hint.</param>
    /// <exception cref="ArgumentException">
    /// Not a colour; a theme slot in a layout without a theme part; a percentage outside 1-99; or a lighter fill.
    /// </exception>
    internal static LayoutColor Parse(string argument, string what, string paramName, WordprocessingDocument doc, ColorUse use)
    {
        ArgumentNullException.ThrowIfNull(argument);
        ArgumentNullException.ThrowIfNull(doc);

        var trimmed = argument.Trim();
        if (string.Equals(trimmed, "auto", StringComparison.OrdinalIgnoreCase))
        {
            return new LayoutColor("auto", null, null, null, null);
        }

        var hex = trimmed.StartsWith('#') ? trimmed[1..] : trimmed;
        if (hex.Length == 6 && hex.All(Uri.IsHexDigit))
        {
            return new LayoutColor(hex.ToUpperInvariant(), null, null, null, null);
        }

        var match = Shape.Match(trimmed);
        var slot = match.Success ? SlotName(match.Groups["slot"].Value) : null;
        if (slot is null)
        {
            throw new ArgumentException(
                $"{what} must be six hex digits RRGGBB (e.g. '0070C0'; a leading '#' is fine), 'auto', or a theme colour "
                + $"({string.Join(", ", ThemePalette.Slots)}), optionally 'lighter N%' or 'darker N%' as Word's palette "
                + $"names them, e.g. 'accent1' or 'text2 lighter 40%' (got '{argument}').",
                paramName);
        }

        string? tint = null;
        string? shade = null;
        int? percent = null;
        if (match.Groups["direction"].Success)
        {
            percent = int.Parse(match.Groups["percent"].Value, CultureInfo.InvariantCulture);
            if (percent is < 1 or > 99)
            {
                throw new ArgumentException(
                    $"{what} '{argument}': the percentage must be between 1 and 99, as in Word's palette ('lighter 40%', "
                    + "'darker 25%').",
                    paramName);
            }

            var lighter = match.Groups["direction"].Value.Equals("lighter", StringComparison.OrdinalIgnoreCase);
            if (lighter && use == ColorUse.Fill)
            {
                throw new ArgumentException(
                    $"{what} '{argument}': a lighter theme fill is not written yet. No captured Business Central layout "
                    + "carries one (w:themeFillTint), and this server only writes shapes a real layout has shown "
                    + $"(ADR-0005). Use '{slot}' or '{slot} darker N%', or the colour as RRGGBB.",
                    paramName);
            }

            var fraction = Math.Round(255 * (100 - percent.Value) / 100.0, MidpointRounding.AwayFromZero);
            var hexByte = ((int)fraction).ToString("X2", CultureInfo.InvariantCulture);
            (tint, shade) = lighter ? (hexByte, (string?)null) : (null, hexByte);
        }

        var main = doc.MainDocumentPart ?? throw new InvalidDataException("Layout has no main document part.");
        if (main.ThemePart is null)
        {
            throw new ArgumentException(
                $"{what} '{argument}' names a theme colour, but this layout has no theme part (word/theme/theme1.xml) to "
                + "take it from, so the colour would follow nothing. A blank create_layout build has none: give the "
                + "colour as RRGGBB, or start from a templatePath whose theme defines the colours.",
                paramName);
        }

        var rgb = ThemePalette.Resolve(main, slot)
            ?? throw new ArgumentException(
                $"{what} '{argument}': the layout's theme defines no colour for '{slot}', so the colour would follow "
                + "nothing. get_layout_info's look.theme lists the slots it defines.",
                paramName);
        return new LayoutColor(ThemePalette.ApplyTintShade(rgb, tint, shade), slot, tint, shade, percent);
    }

    /// <summary>"0070C0", "auto", "accent1 (F24F4F)", "text2 lighter 40% (9B9482)": for a tool summary.</summary>
    internal string Describe()
    {
        if (ThemeSlot is null)
        {
            return Value;
        }

        var variant = Percent is { } percent
            ? string.Create(CultureInfo.InvariantCulture, $" {(Tint is null ? "darker" : "lighter")} {percent}%")
            : string.Empty;
        return $"theme colour {ThemeSlot}{variant} ({Value} in this layout's theme)";
    }

    /// <summary>The run, control or paragraph-mark colour element.</summary>
    internal Color ToRunColor()
    {
        var color = new Color { Val = Value };
        if (ThemeSlot is not null)
        {
            color.ThemeColor = SlotValue(ThemeSlot);
            color.ThemeTint = Attribute(Tint);
            color.ThemeShade = Attribute(Shade);
        }

        return color;
    }

    /// <summary>The cell shading element, in the corpus shape <c>w:val="clear" w:color="auto"</c>.</summary>
    internal Shading ToShading()
    {
        var shading = new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = Value };
        if (ThemeSlot is not null)
        {
            shading.ThemeFill = SlotValue(ThemeSlot);
            shading.ThemeFillShade = Attribute(Shade);
        }

        return shading;
    }

    /// <summary>Sets <paramref name="edge"/>'s colour, replacing whatever colour attributes it had.</summary>
    internal void ApplyTo(BorderType edge)
    {
        ArgumentNullException.ThrowIfNull(edge);

        edge.Color = Value;
        edge.ThemeColor = ThemeSlot is null ? null : SlotValue(ThemeSlot);
        edge.ThemeTint = Attribute(Tint);
        edge.ThemeShade = Attribute(Shade);
    }

    // A null string converts to an EMPTY attribute value, which the schema refuses; null removes the attribute.
    private static StringValue? Attribute(string? value) => value is null ? null : new StringValue(value);

    /// <summary>The slot as <c>w:themeColor</c> spells it (case-insensitive input), or null when it names none.</summary>
    private static string? SlotName(string name) =>
        ThemePalette.Slots.FirstOrDefault(s => s.Equals(name, StringComparison.OrdinalIgnoreCase));

    private static ThemeColorValues SlotValue(string slot) => slot switch
    {
        "text1" => ThemeColorValues.Text1,
        "background1" => ThemeColorValues.Background1,
        "text2" => ThemeColorValues.Text2,
        "background2" => ThemeColorValues.Background2,
        "accent1" => ThemeColorValues.Accent1,
        "accent2" => ThemeColorValues.Accent2,
        "accent3" => ThemeColorValues.Accent3,
        "accent4" => ThemeColorValues.Accent4,
        "accent5" => ThemeColorValues.Accent5,
        "accent6" => ThemeColorValues.Accent6,
        "hyperlink" => ThemeColorValues.Hyperlink,
        _ => ThemeColorValues.FollowedHyperlink,
    };
}
