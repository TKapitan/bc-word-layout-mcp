using System.ComponentModel;
using BcWordLayout.Domain.Models;

namespace BcWordLayout.McpHost.Tools;

/// <summary>
/// The <c>format</c> argument every text-creating tool and <c>set_text_format</c> take: one object with the same
/// members and meaning everywhere (GitHub issue #15), mapped onto the domain's <see cref="TextFormat"/>. Its
/// member descriptions are what an agent reads in the tool schema, so they are written once, here, instead of
/// once per tool. Unknown member names are refused at the call boundary like unknown argument names
/// (<c>ToolCallBoundary</c>), and the whole object may also arrive JSON-encoded as a string.
/// </summary>
public sealed record TextFormatArgument
{
    /// <summary>The <c>format</c> parameter's description on the four insert tools.</summary>
    internal const string InsertDescription =
        "Optional formatting for the inserted text, as one object, e.g. {\"bold\": true, \"fontSizePoints\": 8}: "
        + "bold, italic, underline, fontSizePoints, fontName, color, characterStyle. A member you omit keeps the "
        + "formatting the text takes from the text before it - the anchor control for afterControl, the "
        + "paragraph's last text for tableCell/atText, or an empty paragraph's own mark. Text that opens a new "
        + "paragraph, or lands in a fresh plain-table cell, has nothing to take and is unstyled.";

    /// <summary>The <c>format</c> parameter's description on <c>set_cell_text</c>.</summary>
    internal const string SetCellTextDescription =
        "Optional formatting for the new text, the same object the insert tools take (bold, italic, underline, "
        + "fontSizePoints, fontName, color, characterStyle), e.g. {\"bold\": true}. A member you omit keeps the "
        + "formatting of the text it replaces.";

    /// <summary>The <c>format</c> parameter's description on the two row tools.</summary>
    internal const string RowDescription =
        "Optional formatting for every bound cell's text in the new row, the same object the insert tools take "
        + "(bold, italic, underline, fontSizePoints, fontName, color, characterStyle), e.g. {\"bold\": true} for "
        + "the stock total look. Omit it to leave the controls unstyled.";

    [Description("true makes the text bold; false makes it not bold, even where a character style such as "
                 + "'Strong' supplies the bold.")]
    public bool? Bold { get; init; }

    [Description("true makes the text italic; false makes it not italic, even where a character style supplies it.")]
    public bool? Italic { get; init; }

    [Description("true underlines the text with a single line; false removes the underline.")]
    public bool? Underline { get; init; }

    [Description("Font size in points, 4 to 96; halves allowed (8.5).")]
    public double? FontSizePoints { get; init; }

    [Description("Font family name, e.g. 'Segoe UI' (what Microsoft's own layouts use).")]
    public string? FontName { get; init; }

    [Description("Text colour as six hex digits RRGGBB, e.g. '0070C0' (a leading '#' is fine), or 'auto'.")]
    public string? Color { get; init; }

    [Description("Style id of a CHARACTER style the layout's styles part defines, e.g. 'Strong'; refused when the "
                 + "layout does not define it (the error lists the ones it does).")]
    public string? CharacterStyle { get; init; }

    /// <summary>The domain record, or null when the argument is absent or sets nothing.</summary>
    internal static TextFormat? ToDomain(TextFormatArgument? format)
    {
        if (format is null)
        {
            return null;
        }

        var domain = new TextFormat
        {
            Bold = format.Bold,
            Italic = format.Italic,
            Underline = format.Underline,
            FontSizePoints = format.FontSizePoints,
            FontName = format.FontName,
            Color = format.Color,
            CharacterStyle = format.CharacterStyle,
        };
        return domain.IsEmpty ? null : domain;
    }
}
