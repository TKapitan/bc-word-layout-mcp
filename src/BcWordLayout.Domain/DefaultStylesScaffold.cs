using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Domain;

/// <summary>
/// Creates the default <see cref="StyleDefinitionsPart"/> a BLANK <see cref="LayoutBuilder.Create"/> output
/// needs before its typography means anything. A from-scratch layout used to ship no <c>word/styles.xml</c>
/// and no theme at all, so nothing in the file named a typeface — Word rendered its application default and
/// Business Central rendered a DIFFERENT one (observed against a real BC sandbox, 2026-08-01; GitHub issue
/// #3): same document, two fonts, and no way for a caller to control it short of hand-editing OOXML or
/// supplying a <c>templatePath</c>. The scaffold pins the typography every stock corpus layout resolves to —
/// Calibri 11&#160;pt — so a blank build renders identically on both sides, exactly like
/// <see cref="HeaderFooterScaffold"/> pins the corpus header/footer shape and
/// <see cref="LayoutBuilder"/>'s own trailing <c>w:sectPr</c> pins the corpus page setup.
/// </summary>
/// <remarks>
/// <para>
/// STYLES ONLY, NO THEME PART. The corpus's own <c>docDefaults</c> reach their font through theme tokens
/// (<c>w:asciiTheme="minorHAnsi"</c> → <c>word/theme/theme1.xml</c> → Calibri); this scaffold names Calibri
/// DIRECTLY in <c>w:rFonts</c> instead, which makes the theme part pure dead weight (~7&#160;KB of
/// clrScheme/fmtScheme XML with no rendering effect) and it is deliberately not emitted. Every element shape
/// used here is corpus-observed per the observed-OOXML-only rule: explicit-font <c>w:rFonts</c> mirrors
/// <c>StandardPurchaseOrder.docx</c>'s <c>Style1</c> attribute-for-attribute (including its
/// <c>w:cs="Times New Roman"</c>), the <c>w:sz</c>/<c>w:szCs</c> values and <c>w:pPrDefault</c> spacing
/// mirror the corpus <c>docDefaults</c>, and each named style below mirrors
/// <c>StandardPurchaseOrder.docx</c>'s definition element-for-element (minus <c>w:rsid</c>, Word's own
/// revision-save bookkeeping, which this tool never emits).
/// </para>
/// <para>
/// WHAT IS EMITTED: <c>docDefaults</c> plus Word's four stock default styles — <c>Normal</c>,
/// <c>DefaultParagraphFont</c>, <c>TableNormal</c>, <c>NoList</c> — plus <c>TableGrid</c>, so
/// <c>insert_repeater_table</c>'s documented <c>tableStyle='TableGrid'</c> example resolves on a
/// from-scratch layout instead of writing a <c>w:tblStyle</c> reference that points at nothing (issue #3's
/// knock-on; <c>LayoutValidator</c>'s <c>table-style-resolves</c> check warns about any that still don't).
/// No <c>latentStyles</c> — it only affects Word's style-gallery UI, never rendering.
/// </para>
/// <para>
/// BLANK BUILDS ONLY, AND ONLY AT CREATE TIME. A <c>templatePath</c> brings its own look (styles, theme,
/// or deliberately neither) and must never have parts injected into it — same contract as
/// <see cref="HeaderFooterScaffold"/> on the template path. Unlike that scaffold, this one is also NEVER
/// applied on demand by <see cref="LayoutEditor"/> to a pre-existing layout: an existing document already
/// renders SOMEHOW, and retrofitting <c>docDefaults</c> would silently change how it looks everywhere. The
/// only legitimate moment to pin a default is the moment the document is born with no look to preserve.
/// </para>
/// </remarks>
public static class DefaultStylesScaffold
{
    /// <summary>The typeface the scaffold pins — what every stock corpus layout resolves <c>minorHAnsi</c> to.</summary>
    internal const string DefaultFont = "Calibri";

    /// <summary>Half-point font size the corpus <c>docDefaults</c> carry (22 = 11&#160;pt).</summary>
    internal const string DefaultFontSizeHalfPoints = "22";

    /// <summary>
    /// Ensures <paramref name="main"/> has a <see cref="StyleDefinitionsPart"/>, adding the default one
    /// described in this type's remarks when it has none. Returns <c>true</c> only when the part was
    /// actually added — a document that already has one (e.g. any real corpus layout) is a no-op returning
    /// <c>false</c>, its own styles left byte-for-byte untouched.
    /// </summary>
    public static bool EnsureDefaultStyles(MainDocumentPart main)
    {
        ArgumentNullException.ThrowIfNull(main);
        if (main.StyleDefinitionsPart is not null)
        {
            return false;
        }

        var stylesPart = main.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new Styles(
            BuildDocDefaults(),
            BuildNormalStyle(),
            BuildDefaultParagraphFontStyle(),
            BuildTableNormalStyle(),
            BuildNoListStyle(),
            BuildTableGridStyle());
        stylesPart.Styles.Save();
        return true;
    }

    /// <summary>
    /// <c>w:docDefaults</c>: Calibri 11&#160;pt named explicitly (no theme indirection — see the type
    /// remarks) plus the corpus's own default paragraph spacing (<c>after=160</c>, <c>line=259 auto</c>).
    /// The explicit <c>w:rFonts</c> attribute set mirrors <c>StandardPurchaseOrder.docx</c>'s
    /// <c>Style1</c>; the sizes and spacing mirror the corpus <c>docDefaults</c> themselves.
    /// </summary>
    private static DocDefaults BuildDocDefaults() => new(
        new RunPropertiesDefault(new RunPropertiesBaseStyle(
            new RunFonts
            {
                Ascii = DefaultFont,
                EastAsia = DefaultFont,
                HighAnsi = DefaultFont,
                ComplexScript = "Times New Roman",
            },
            new FontSize { Val = DefaultFontSizeHalfPoints },
            new FontSizeComplexScript { Val = DefaultFontSizeHalfPoints })),
        new ParagraphPropertiesDefault(new ParagraphPropertiesBaseStyle(
            new SpacingBetweenLines { After = "160", Line = "259", LineRule = LineSpacingRuleValues.Auto })));

    private static Style BuildNormalStyle() => new(
        new StyleName { Val = "Normal" },
        new PrimaryStyle(),
        new StyleParagraphProperties(
            new SpacingBetweenLines { After = "200", Line = "240", LineRule = LineSpacingRuleValues.Auto }))
    {
        Type = StyleValues.Paragraph,
        Default = true,
        StyleId = "Normal",
    };

    private static Style BuildDefaultParagraphFontStyle() => new(
        new StyleName { Val = "Default Paragraph Font" },
        new UIPriority { Val = 1 },
        new SemiHidden(),
        new UnhideWhenUsed())
    {
        Type = StyleValues.Character,
        Default = true,
        StyleId = "DefaultParagraphFont",
    };

    private static Style BuildTableNormalStyle() => new(
        new StyleName { Val = "Normal Table" },
        new UIPriority { Val = 99 },
        new SemiHidden(),
        new UnhideWhenUsed(),
        new StyleTableProperties(
            new TableIndentation { Width = 0, Type = TableWidthUnitValues.Dxa },
            new TableCellMarginDefault(
                new TopMargin { Width = "0", Type = TableWidthUnitValues.Dxa },
                new TableCellLeftMargin { Width = 108, Type = TableWidthValues.Dxa },
                new BottomMargin { Width = "0", Type = TableWidthUnitValues.Dxa },
                new TableCellRightMargin { Width = 108, Type = TableWidthValues.Dxa })))
    {
        Type = StyleValues.Table,
        Default = true,
        StyleId = "TableNormal",
    };

    private static Style BuildNoListStyle() => new(
        new StyleName { Val = "No List" },
        new UIPriority { Val = 99 },
        new SemiHidden(),
        new UnhideWhenUsed())
    {
        Type = StyleValues.Numbering,
        Default = true,
        StyleId = "NoList",
    };

    /// <summary>
    /// <c>TableGrid</c> — the one non-default style shipped, because it is the documented example for
    /// <c>insert_repeater_table</c>'s <c>tableStyle</c> parameter and every corpus layout defines it.
    /// Single ½&#160;pt borders all round, compact single-spaced cell paragraphs.
    /// </summary>
    private static Style BuildTableGridStyle() => new(
        new StyleName { Val = "Table Grid" },
        new BasedOn { Val = "TableNormal" },
        new UIPriority { Val = 39 },
        new StyleParagraphProperties(
            new SpacingBetweenLines { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }),
        new StyleTableProperties(
            new TableBorders
            {
                TopBorder = new TopBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
                LeftBorder = new LeftBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
                BottomBorder = new BottomBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
                RightBorder = new RightBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
                InsideHorizontalBorder =
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
                InsideVerticalBorder =
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
            }))
    {
        Type = StyleValues.Table,
        StyleId = "TableGrid",
    };
}
