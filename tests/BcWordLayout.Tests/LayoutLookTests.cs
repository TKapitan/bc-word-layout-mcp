using BcWordLayout.Domain;
using BcWordLayout.McpHost;
using BcWordLayout.McpHost.Tools;
using DocumentFormat.OpenXml.Packaging;

namespace BcWordLayout.Tests;

/// <summary>
/// <c>get_layout_info</c>'s <c>look</c> (GitHub issue #149): the theme, the colours the content uses split by kind,
/// carrier and theme slot, and the style ids in use, checked against the corpus captures whose look the issue
/// describes.
/// </summary>
public class LayoutLookTests
{
    private static LayoutLook Look(string corpusFile)
    {
        using var doc = WordprocessingDocument.Open(Corpus.Path(corpusFile), false);
        return LayoutLookReader.Read(doc);
    }

    private static int Count(LayoutLook look, string part, string kind, string carrier, string color, string? themeColor) =>
        look.Colors.SingleOrDefault(c => c.Part == part && c.Kind == kind && c.Carrier == carrier && c.Color == color
                                         && c.ThemeColor == themeColor)?.Count ?? 0;

    [Fact]
    public void The_themable_invoice_colours_its_text_fills_and_rules_through_accent1()
    {
        var look = Look(Corpus.SalesInvoiceBlueSimpleThemable);

        var theme = Assert.IsType<ThemeInfo>(look.Theme);
        Assert.Equal("Business Invoice", theme.ColorScheme);
        Assert.Equal("F24F4F", theme.Colors.Single(c => c.Key == "accent1").Value);
        Assert.Equal("000000", theme.Colors.Single(c => c.Key == "text1").Value);
        Assert.Equal(ThemePalette.Slots, theme.Colors.Select(c => c.Key));
        Assert.Equal("Segoe UI Light", theme.MajorFont);
        Assert.Equal("Segoe UI", theme.MinorFont);

        // 26 theme text colours: body 14 (7 runs, 6 controls, 1 paragraph mark), header 4, footer 8.
        Assert.Equal(7, Count(look, "document.xml", "text", "run", "F24F4F", "accent1"));
        Assert.Equal(6, Count(look, "document.xml", "text", "control", "F24F4F", "accent1"));
        Assert.Equal(1, Count(look, "document.xml", "text", "paragraphMark", "F24F4F", "accent1"));
        Assert.Equal(26, look.Colors.Where(c => c.Kind == "text" && c.ThemeColor == "accent1").Sum(c => c.Count));
        Assert.Equal(8, Count(look, "document.xml", "fill", "cell", "F24F4F", "accent1"));
        Assert.Equal(25, Count(look, "document.xml", "rule", "cell", "F24F4F", "accent1"));
        Assert.Equal(1, Count(look, "document.xml", "rule", "cell", "9B9482", "text2 lighter 40%"));

        // The literal 0070C0 left behind sits only where no text shows: 6 paragraph marks, 2 control properties.
        var literal = look.Colors.Where(c => c.ThemeColor is null).ToList();
        Assert.All(literal, c => Assert.Equal("0070C0", c.Color));
        Assert.Equal(6, literal.Where(c => c.Carrier == "paragraphMark").Sum(c => c.Count));
        Assert.Equal(2, literal.Where(c => c.Carrier == "control").Sum(c => c.Count));
        Assert.DoesNotContain(literal, c => c.Carrier == "run");

        Assert.Contains(look.Styles, s => s is { StyleId: "Heading1", Kind: "paragraph" });
        Assert.Contains(look.Styles, s => s is { StyleId: "Heading2", Kind: "paragraph" });
    }

    [Fact]
    public void The_themable_shipment_shows_the_two_captions_Microsofts_conversion_left_blue()
    {
        // GitHub issue #149's comment: ShipmentDate_Lbl and OrderNo_Lbl kept a literal 0070C0 on their runs, so under a
        // theme those two captions would stay blue.
        var look = Look(Corpus.SalesShipmentBlueThemable);

        Assert.Equal(2, Count(look, "document.xml", "text", "run", "0070C0", null));
        Assert.Equal(26, look.Colors.Where(c => c.Kind == "text" && c.ThemeColor == "accent1").Sum(c => c.Count));
        Assert.Equal(4, look.Colors.Where(c => c.Kind == "fill" && c.ThemeColor == "accent1").Sum(c => c.Count));
    }

    [Fact]
    public void The_plain_invoice_hard_codes_its_blue_and_tints_its_rules()
    {
        var look = Look(Corpus.SalesInvoiceVatSpec);

        Assert.Equal(8, Count(look, "document.xml", "fill", "cell", "0070C0", null));
        Assert.Equal(4, Count(look, "document.xml", "rule", "cell", "002060", null));
        Assert.Equal(9, Count(look, "document.xml", "rule", "cell", "9B9482", "text2 lighter 40%"));
        Assert.Equal(6, Count(look, "document.xml", "text", "run", "0070C0", null));
        Assert.DoesNotContain(look.Colors, c => c.ThemeColor == "accent1" && c.Carrier == "run");
    }

    [Fact]
    public void A_shade_reads_back_in_the_form_the_colour_arguments_take()
    {
        // SalesInvoiceForSubscriptionBilling's rules: background2 with themeShade BF, Word's "Darker 25%".
        var look = Look(Corpus.SalesInvoice);

        Assert.Equal(34, look.Colors.Where(c => c.Kind == "rule" && c.ThemeColor == "background2 darker 25%").Sum(c => c.Count));
    }

    [Fact]
    public void A_layout_without_a_theme_or_colours_has_an_empty_look()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(rows: 1, cols: 2));
        try
        {
            using var doc = WordprocessingDocument.Open(path, false);
            var look = LayoutLookReader.Read(doc);

            Assert.Null(look.Theme);
            Assert.Empty(look.Colors);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void get_layout_info_reports_the_look()
    {
        var response = ReadTools.GetLayoutInfo(Corpus.Path(Corpus.SalesInvoiceBlueSimpleThemable));

        Assert.True(response.Ok, response.Error?.Message);
        var look = Assert.IsType<LayoutInfoDto>(response.Data).Look;
        Assert.Equal("F24F4F", look.Theme!.Colors["accent1"]);
        Assert.Contains(look.Colors, c => c is { Part: "header1.xml", Kind: "text", Carrier: "run", ThemeColor: "accent1", Count: 2 });
        Assert.Contains(look.Styles, s => s is { StyleId: "Heading1", Kind: "paragraph" });
    }
}
