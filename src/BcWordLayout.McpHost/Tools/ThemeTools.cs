using System.ComponentModel;
using BcWordLayout.Domain;
using ModelContextProtocol.Server;
using static BcWordLayout.McpHost.Tools.ToolGuards;

namespace BcWordLayout.McpHost.Tools;

/// <summary>
/// MCP tools for a layout's own theme part (<c>word/theme/theme1.xml</c>): <c>set_layout_theme</c> changes the colours
/// its scheme gives the theme slots and its heading and body fonts (GitHub issue #150). Same open/validate/save-or-reject
/// choreography as every other mutating tool, through <see cref="ToolGuards.GuardThemeEdit"/>.
/// </summary>
[McpServerToolType]
public static class ThemeTools
{
    [McpServerTool(Name = "set_layout_theme")]
    [Description("Change a layout's OWN theme (word/theme/theme1.xml): the colour its scheme gives each theme slot, "
                 + "and its heading and body fonts. Text, cell fills and rules coloured through a theme slot "
                 + "(format.color, fill or borderColor 'accent1', 'text2 lighter 40%') take the new colour, as do the "
                 + "styles that use the slot: the stock themable Blue invoice prints coral red because its theme's "
                 + "accent1 is F24F4F, and colors {\"accent1\": \"0070C0\"} makes it blue. Word caches the colour "
                 + "each theme reference resolves to beside it; every cached value of a changed slot is rewritten too, "
                 + "tints and shades applied, so a renderer that reads only the cache prints the new colour. Text using "
                 + "the theme fonts (headings, and body text in layouts whose styles say minorHAnsi) takes the new "
                 + "fonts; text set in a named font keeps it. A layout with no theme part is refused: create_layout "
                 + "with theme 'office' gives a new layout one. Needs a BC 29 target, like the other themable "
                 + "formatting (bc_version_unsupported below 29 or when the target is unknown). Whether a Business "
                 + "Central report theme overrides a layout's own theme is not yet known (GitHub issue #142). Same "
                 + "write/validate/save-or-reject safety as every editing tool.")]
    public static ToolResponse SetLayoutTheme(
        [Description("Absolute path to the .docx layout file.")] string layoutPath,
        [Description("Theme slots to set, each to six hex digits RRGGBB (a leading '#' is fine), e.g. "
                     + "{\"accent1\": \"0070C0\", \"text2\": \"44546A\"}. Slots: text1, background1, text2, "
                     + "background2, accent1-accent6, hyperlink, followedHyperlink (case-insensitive). Omit to leave "
                     + "the colours alone.")]
        IReadOnlyDictionary<string, string>? colors = null,
        [Description("The theme's heading font (a:majorFont), a font family name such as 'Segoe UI Semibold'. Omit "
                     + "to leave it.")] string? majorFont = null,
        [Description("The theme's body font (a:minorFont), a font family name such as 'Segoe UI'. Omit to leave it.")]
        string? minorFont = null)
    {
        return GuardThemeEdit(layoutPath, (doc, target) =>
        {
            BcFeatures.Require(BcFeatures.ThemableFormatting, target);
            return ThemeEditor.SetTheme(doc, colors, majorFont, minorFont);
        });
    }
}
