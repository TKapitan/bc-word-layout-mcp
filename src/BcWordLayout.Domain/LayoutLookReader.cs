using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Domain;

/// <summary>
/// One group of identical colour references in one part (GitHub issue #149): what is coloured, on what, with which
/// colour, and how often.
/// </summary>
/// <param name="Part">The content part, e.g. <c>document.xml</c>, <c>footer2.xml</c>.</param>
/// <param name="Kind"><c>text</c> (<c>w:color</c>), <c>fill</c> (<c>w:shd</c>) or <c>rule</c> (a drawn border).</param>
/// <param name="Carrier">
/// What carries it: <c>run</c> (the text itself), <c>paragraphMark</c> (a paragraph's mark, <c>w:pPr/w:rPr</c>, which
/// shows no text of its own), <c>control</c> (a content control's own formatting, <c>w:sdtPr/w:rPr</c>), <c>cell</c>,
/// <c>table</c>, <c>paragraph</c> or <c>page</c>.
/// </param>
/// <param name="Color">The literal colour stored (<c>RRGGBB</c>); for a theme colour, the value Word cached for it.</param>
/// <param name="ThemeColor">
/// The theme colour, in the form the colour arguments take (<c>accent1</c>, <c>text2 lighter 40%</c>), or null for a
/// literal colour.
/// </param>
/// <param name="Count">How many elements carry exactly this.</param>
public sealed record ColorUsage(string Part, string Kind, string Carrier, string Color, string? ThemeColor, int Count);

/// <summary>A style id the layout's content refers to, by the kind of reference, and how often.</summary>
/// <param name="StyleId">The style id, e.g. <c>Heading1</c>.</param>
/// <param name="Kind"><c>paragraph</c> (<c>w:pStyle</c>), <c>character</c> (<c>w:rStyle</c>) or <c>table</c> (<c>w:tblStyle</c>).</param>
/// <param name="Count">How many references, across the body, headers and footers.</param>
public sealed record StyleUsage(string StyleId, string Kind, int Count);

/// <summary>The look of a layout: its theme, the colours its content uses, and the styles it refers to.</summary>
public sealed record LayoutLook(ThemeInfo? Theme, IReadOnlyList<ColorUsage> Colors, IReadOnlyList<StyleUsage> Styles);

/// <summary>
/// Reads what <c>get_layout_info</c> reports as <c>look</c> (GitHub issue #149), so an agent can tell whether a layout
/// colours through its theme, or hard-codes a colour, without reading the OOXML: before restyling it, keeping it
/// themable, or checking the style contract between a body and a theme.
/// </summary>
/// <remarks>
/// <para>
/// Only the content parts (body, headers, footers) are read: style definitions colour only the text that uses them,
/// which <see cref="LayoutLook.Styles"/> names. Recorded revisions (<c>w:rPrChange</c> and the like) describe a past
/// look and are skipped. <c>auto</c> colours, no-fill shading and cleared edges are not colours anyone chose, so they
/// are left out.
/// </para>
/// <para>
/// The carrier matters for a themable layout. Microsoft's own themable captures keep literal <c>0070C0</c> on
/// paragraph marks and on a control's own formatting while every run follows <c>accent1</c>
/// (StandardSalesInvoiceBlueSimpleThemable), so a literal colour there does not by itself mean the text is
/// hard-coded; StandardSalesShipmentBlueThemable has two captions whose RUNS kept <c>0070C0</c>, which is.
/// </para>
/// </remarks>
public static class LayoutLookReader
{
    /// <summary>Reads <paramref name="doc"/>'s look.</summary>
    public static LayoutLook Read(WordprocessingDocument doc)
    {
        ArgumentNullException.ThrowIfNull(doc);

        var main = doc.MainDocumentPart ?? throw new InvalidDataException("Layout has no main document part.");
        var colors = new List<ColorUsage>();
        var styles = new Dictionary<(string Id, string Kind), int>();
        var styleOrder = new List<(string Id, string Kind)>();

        foreach (var (root, partName) in PartWalker.ContentParts(main))
        {
            var groups = new Dictionary<(string Kind, string Carrier, string Color, string? Theme), int>();
            void Add(string kind, string carrier, string color, string? theme)
            {
                var key = (kind, carrier, color, theme);
                groups[key] = groups.TryGetValue(key, out var n) ? n + 1 : 1;
            }

            foreach (var element in root.Descendants())
            {
                if (IsRevisionHistory(element))
                {
                    continue;
                }

                switch (element)
                {
                    case Color color when TextCarrier(color) is { } carrier:
                        var theme = ThemeArgument(color.ThemeColor?.InnerText, color.ThemeTint?.Value, color.ThemeShade?.Value);
                        if (theme is not null || IsChosen(color.Val?.Value))
                        {
                            Add("text", carrier, color.Val?.Value ?? "auto", theme);
                        }

                        break;

                    case Shading shading:
                        var fillTheme = ThemeArgument(
                            shading.ThemeFill?.InnerText, shading.ThemeFillTint?.Value, shading.ThemeFillShade?.Value);
                        if (fillTheme is not null || IsChosen(shading.Fill?.Value))
                        {
                            Add("fill", Carrier(shading.Parent), shading.Fill?.Value ?? "auto", fillTheme);
                        }

                        break;

                    case BorderType edge when edge.Val?.Value is { } line && line != BorderValues.Nil && line != BorderValues.None:
                        var ruleTheme = ThemeArgument(edge.ThemeColor?.InnerText, edge.ThemeTint?.Value, edge.ThemeShade?.Value);
                        if (ruleTheme is not null || IsChosen(edge.Color?.Value))
                        {
                            Add("rule", Carrier(edge.Parent), edge.Color?.Value ?? "auto", ruleTheme);
                        }

                        break;

                    case ParagraphStyleId pStyle:
                        Count(pStyle.Val?.Value, "paragraph");
                        break;

                    case RunStyle rStyle:
                        Count(rStyle.Val?.Value, "character");
                        break;

                    case TableStyle tblStyle:
                        Count(tblStyle.Val?.Value, "table");
                        break;
                }
            }

            colors.AddRange(groups
                .OrderBy(g => KindOrder(g.Key.Kind))
                .ThenByDescending(g => g.Value)
                .ThenBy(g => g.Key.Carrier, StringComparer.Ordinal)
                .ThenBy(g => g.Key.Color, StringComparer.Ordinal)
                .Select(g => new ColorUsage(partName, g.Key.Kind, g.Key.Carrier, g.Key.Color, g.Key.Theme, g.Value)));
        }

        return new LayoutLook(
            ThemePalette.Read(main),
            colors,
            styleOrder.Select(k => new StyleUsage(k.Id, k.Kind, styles[k])).ToList());

        void Count(string? styleId, string kind)
        {
            if (string.IsNullOrEmpty(styleId))
            {
                return;
            }

            var key = (styleId, kind);
            if (styles.TryGetValue(key, out var n))
            {
                styles[key] = n + 1;
            }
            else
            {
                styles[key] = 1;
                styleOrder.Add(key);
            }
        }
    }

    /// <summary>
    /// A theme colour in the form the colour arguments take: the slot, plus <c>lighter N%</c> / <c>darker N%</c> for a
    /// tint or shade byte (see <see cref="LayoutColor"/>); null when <paramref name="slot"/> is absent or <c>none</c>.
    /// </summary>
    internal static string? ThemeArgument(string? slot, string? tint, string? shade)
    {
        if (string.IsNullOrEmpty(slot) || slot == "none")
        {
            return null;
        }

        return tint is not null ? $"{slot} lighter {Percent(tint)}%"
            : shade is not null ? $"{slot} darker {Percent(shade)}%"
            : slot;
    }

    private static string Percent(string hexByte) =>
        int.TryParse(hexByte, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var value)
            ? Math.Round(100 - (value * 100 / 255.0), MidpointRounding.AwayFromZero).ToString(CultureInfo.InvariantCulture)
            : hexByte;

    private static bool IsChosen(string? value) =>
        !string.IsNullOrEmpty(value) && !string.Equals(value, "auto", StringComparison.OrdinalIgnoreCase);

    private static int KindOrder(string kind) => kind switch
    {
        "text" => 0,
        "fill" => 1,
        _ => 2,
    };

    /// <summary>What carries a text colour, or null for a <c>w:color</c> outside run properties (none in a content part).</summary>
    private static string? TextCarrier(Color color) => color.Parent switch
    {
        ParagraphMarkRunProperties => "paragraphMark",
        RunProperties { Parent: Run } => "run",
        RunProperties { Parent: SdtProperties or SdtEndCharProperties } => "control",
        _ => null,
    };

    /// <summary>What carries a fill or a rule, from the properties element it sits in.</summary>
    private static string Carrier(OpenXmlElement? parent) => parent switch
    {
        TableCellProperties or TableCellBorders => "cell",
        TableProperties or TableBorders or TablePropertyExceptions => "table",
        ParagraphProperties or ParagraphBorders => "paragraph",
        PageBorders => "page",
        ParagraphMarkRunProperties => "paragraphMark",
        RunProperties { Parent: SdtProperties or SdtEndCharProperties } => "control",
        _ => "run",
    };

    /// <summary>True inside a recorded revision (<c>w:rPrChange</c>, <c>w:pPrChange</c>, ...): a past look, not this one.</summary>
    private static bool IsRevisionHistory(OpenXmlElement element) =>
        element.Ancestors().Any(a => a is RunPropertiesChange or ParagraphPropertiesChange or TableCellPropertiesChange
            or TablePropertiesChange or TableRowPropertiesChange or SectionPropertiesChange or TablePropertyExceptionsChange
            or ParagraphMarkRunPropertiesChange);
}
