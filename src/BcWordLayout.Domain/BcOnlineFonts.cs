using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Domain;

/// <summary>
/// The fonts Business Central online has installed for rendering report layouts, and whether a font a layout
/// names is one of them or is embedded in the layout.
/// </summary>
/// <remarks>
/// <para>
/// WHY: Business Central prints text in a font it lacks in a substitute. Sandbox probe P13 (BC 28, 2026-10-24)
/// printed a line set in a font that exists nowhere in a serif face, while the same probe's Georgia and Segoe UI
/// lines printed as set. Microsoft documents the fallback as the Normal style's font, and custom fonts cannot be
/// uploaded to BC online; a font embedded in the layout is used.
/// </para>
/// <para>
/// THE LIST is Microsoft's "Available fonts" page (learn.microsoft.com/dynamics365/business-central/ui-fonts,
/// "Document fonts", read 2026-10-25), 519 names including the barcode, OCR and MICR fonts. Matching ignores
/// case, as Word does. The page cuts some names to 31 characters, the limit of a legacy font family name
/// ('Bahnschrift SemiBold SemiConden'), so a 31-character entry matches every name it begins; it also lists some
/// barcode fonts by file name, so a trailing '.ttf' is ignored. An on-premises server renders with the fonts
/// installed on it, which this list does not describe.
/// </para>
/// </remarks>
internal static class BcOnlineFonts
{
    private const int LegacyFamilyNameLength = 31;

    private static readonly string[] Listed =
    [
        "Aharoni", "Aldhabi", "Andalus", "Angsana New", "AngsanaUPC", "Aparajita", "Aptos", "Aptos Black",
        "Aptos Display", "Aptos ExtraBold", "Aptos Light", "Aptos Mono", "Aptos Narrow", "Aptos SemiBold",
        "Aptos Serif", "Arabic Typesetting", "Arial", "Arial Black", "Arial Narrow", "Arial Nova", "Arial Nova Cond",
        "Arial Nova Cond Light", "Arial Nova Light", "Bahnschrift", "Bahnschrift Condensed", "Bahnschrift Light",
        "Bahnschrift Light Condensed", "Bahnschrift Light SemiCondensed", "Bahnschrift SemiBold",
        "Bahnschrift SemiBold Condensed", "Bahnschrift SemiBold SemiConden", "Bahnschrift SemiCondensed",
        "Bahnschrift SemiLight", "Bahnschrift SemiLight Condensed", "Bahnschrift SemiLight SemiConde", "Batang",
        "BatangChe", "BIZ UDGothic", "BIZ UDMincho Medium", "BIZ UDPGothic", "BIZ UDPMincho Medium", "Browallia New",
        "BrowalliaUPC", "Calibri", "Calibri Light", "Cambria", "Cambria Math", "Candara", "Candara Light",
        "Cascadia Code", "Cascadia Code ExtraLight", "Cascadia Code Light", "Cascadia Code SemiBold",
        "Cascadia Code SemiLight", "Cascadia Mono", "Cascadia Mono ExtraLight", "Cascadia Mono Light",
        "Cascadia Mono SemiBold", "Cascadia Mono SemiLight", "Comic Sans MS", "Consolas", "Constantia", "Corbel",
        "Corbel Light", "Cordia New", "CordiaUPC", "Courier New", "DaunPenh", "David", "DengXian", "DengXian Light",
        "DFKai-SB", "DilleniaUPC", "DokChampa", "Dotum", "DotumChe", "E-13B", "Ebrima", "Estrangelo Edessa",
        "EucrosiaUPC", "Euphemia", "FangSong", "Franklin Gothic Medium", "Franklin Gothic Medium Cond", "FrankRuehl",
        "FreesiaUPC", "Gabriola", "Gadugi", "Gautami", "Georgia", "Georgia Pro", "Georgia Pro Black",
        "Georgia Pro Cond", "Georgia Pro Cond Black", "Georgia Pro Cond Light", "Georgia Pro Cond Semibold",
        "Georgia Pro Light", "Georgia Pro Semibold", "Gill Sans Nova", "Gill Sans Nova Cond",
        "Gill Sans Nova Cond Lt", "Gill Sans Nova Cond Ultra Bold", "Gill Sans Nova Cond XBd",
        "Gill Sans Nova Light", "Gill Sans Nova Ultra Bold", "Gisha", "Gulim", "GulimChe", "Gungsuh", "GungsuhChe",
        "HoloLens MDL2 Assets", "IDAutomation Maxicode", "IDAutomation Symbol Text", "IDAutomation2D",
        "IDAutomation4State", "IDAutomation4STATEn1", "IDAutomationBHI25L", "IDAutomationBHI25M",
        "IDAutomationBHI25S", "IDAutomationBHI25XL", "IDAutomationBHI25XS", "IDAutomationBHI25XXL",
        "IDAutomationBI25L", "IDAutomationBI25M", "IDAutomationBI25S", "IDAutomationBI25XL", "IDAutomationBI25XS",
        "IDAutomationBI25XXL", "IDAutomationC11L.ttf", "IDAutomationC11M.ttf", "IDAutomationC11S.ttf",
        "IDAutomationC11XS.ttf", "IDAutomationC128L", "IDAutomationC128M", "IDAutomationC128S", "IDAutomationC128XL",
        "IDAutomationC128XS", "IDAutomationC128XXL", "IDAutomationC128XXS", "IDAutomationC128L SymbolEncoded.ttf",
        "IDAutomationC128M SymbolEncoded.ttf", "IDAutomationC128S SymbolEncoded.ttf",
        "IDAutomationC128XS SymbolEncoded.ttf", "IDAutomationC128XXS SymbolEncoded.ttf", "IDAutomationC25L.ttf",
        "IDAutomationC25M.ttf", "IDAutomationC25S.ttf", "IDAutomationC25XS.ttf", "IDAutomationC39L",
        "IDAutomationC39M", "IDAutomationC39S", "IDAutomationC39XL", "IDAutomationC39XS", "IDAutomationC39XXL",
        "IDAutomationC39n240m.ttf", "IDAutomationC39n240s.ttf", "IDAutomationC39n240xs.ttf",
        "IDAutomationC39n240xxs.ttf", "IDAutomationC39nL.ttf", "IDAutomationC39nM.ttf", "IDAutomationC39nS.ttf",
        "IDAutomationC39nXL.ttf", "IDAutomationC39nXS.ttf", "IDAutomationHC39nL.ttf", "IDAutomationHC39nM.ttf",
        "IDAutomationHC39nS.ttf", "IDAutomationHC39nXL.ttf", "IDAutomationHC39nXS.ttf", "IDAutomationC93L",
        "IDAutomationC93M", "IDAutomationC93S", "IDAutomationC93XL", "IDAutomationC93XS", "IDAutomationCBL",
        "IDAutomationCBM", "IDAutomationCBS", "IDAutomationCBXL", "IDAutomationCBXS", "IDAutomationCBXXL",
        "IDAutomationCMC7", "IDAutomationCMC7n10", "IDAutomationCMC7n25", "IDAutomationCMC7n40", "IDAutomationFIM",
        "IDAutomationHbC128L", "IDAutomationHbC128M", "IDAutomationHbC128S", "IDAutomationHbC128XL",
        "IDAutomationHbC128XS", "IDAutomationHbC128XXL", "IDAutomationHC39L", "IDAutomationHC39M",
        "IDAutomationHC39S", "IDAutomationHC39XL", "IDAutomationHC39XS", "IDAutomationHC39XXL", "IDAutomationHCBL",
        "IDAutomationHCBM", "IDAutomationHCBS", "IDAutomationHCBXL", "IDAutomationHCBXS", "IDAutomationHCBXXL",
        "IDAutomationHcC128L", "IDAutomationHcC128M", "IDAutomationHcC128S", "IDAutomationHcC128XL",
        "IDAutomationHcC128XS", "IDAutomationHcC128XXL", "IDAutomationHI25L", "IDAutomationHI25M",
        "IDAutomationHI25S", "IDAutomationHI25XL", "IDAutomationHI25XS", "IDAutomationHI25XXL", "IDAutomationHMSIL",
        "IDAutomationHMSIM", "IDAutomationHMSIS", "IDAutomationHMSIXS", "IDAutomationHSec1B", "IDAutomationHSec1W",
        "IDAutomationI25L", "IDAutomationI25M", "IDAutomationI25S", "IDAutomationI25XL", "IDAutomationI25XS",
        "IDAutomationI25XXL", "IDAutomationIMB", "IDAutomationIMBn1", "IDAutomationIMBn2", "IDAutomationMICR",
        "IDAutomationMICR B", "IDAutomationMICR L", "IDAutomationMICR N1", "IDAutomationMICR N1B",
        "IDAutomationMICR N1L", "IDAutomationMICR N1XB", "IDAutomationMICR N1XL", "IDAutomationMICR N2",
        "IDAutomationMICR N2B", "IDAutomationMICR N2L", "IDAutomationMICR N2XB", "IDAutomationMICR N2XL",
        "IDAutomationMICR W1", "IDAutomationMICR W1B", "IDAutomationMICR W1L", "IDAutomationMICR W1XB",
        "IDAutomationMICR W1XL", "IDAutomationMICR W2", "IDAutomationMICR W2B", "IDAutomationMICR W2L",
        "IDAutomationMICR W2XB", "IDAutomationMICR W2XL", "IDAutomationMICR XB", "IDAutomationMICR XL",
        "IDAutomationMSIL", "IDAutomationMSIM", "IDAutomationMSIS", "IDAutomationMSIXS", "IDAutomationOCRa",
        "IDAutomationOCRa6LPI.ttf", "IDAutomationOCRaIII.ttf", "IDAutomationOCRaIV.ttf", "IDAutomationOCRbIII.ttf",
        "IDAutomationOCRbIV.ttf", "IDAutomationPLANET", "IDAutomationPLANETn1", "IDAutomationPLANETn2",
        "IDAutomationPOSTNET", "IDAutomationPOSTNETn1", "IDAutomationPOSTNETn2", "IDAutomationRM",
        "IDAutomationRMn1", "IDAutomationSec1B", "IDAutomationSec1W", "IDAutomationSecurityBlock",
        "IDAutomationUPCEANL", "IDAutomationUPCEANM", "IDAutomationUPCEANMn", "IDAutomationUPCEANS",
        "IDAutomationUPCEANSn", "IDAutomationUPCEANXS", "IDAutomationUPCEANXSn", "IDAutomationUPCEANXXS",
        "IDAutomationUPCEANXXSn", "IDAutomationUPCEANL Symbol.ttf", "IDAutomationUPCEANM Symbol.ttf",
        "IDAutomationUPCEANMne.ttf", "IDAutomationUPCEANS Symbol.ttf", "IDAutomationUPCEANSne.ttf",
        "IDAutomationUPCEANXS Symbol.ttf", "IDAutomationUPCEANXSne.ttf", "IDAutomationUPCEANXXS Symbol.ttf",
        "IDAutomationUPCEANXXSne.ttf", "IDAutomationVBHI25L.ttf", "IDAutomationVBHI25M.ttf",
        "IDAutomationVBHI25XL.ttf", "IDAutomationXC39L", "IDAutomationXC39M", "IDAutomationXC39S",
        "IDAutomationXC39XL", "IDAutomationXC39XS", "IDAutomationXC39XXL", "IDAutomationXHC39L",
        "IDAutomationXHC39M", "IDAutomationXHC39S", "IDAutomationXHC39XL", "IDAutomationXHC39XS",
        "IDAutomationXHC39XXL", "IDAutomationYBHI25L Symbol", "IDAutomationYBHI25M Symbol",
        "IDAutomationYBHI25S Symbol", "IDAutomationYBHI25XL Symbol Enc", "IDAutomationYBHI25XS Symbol",
        "IDAutomationYBI25L Symbol", "IDAutomationYBI25M Symbol", "IDAutomationYBI25S Symbol",
        "IDAutomationYBI25XL Symbol Enc", "IDAutomationYBI25XS Symbol", "IDAutomationYC39L Symbol",
        "IDAutomationYC39M Symbol", "IDAutomationYC39S Symbol", "IDAutomationYC39XL Symbol Enc",
        "IDAutomationYC39XS Symbol", "IDAutomationYC39XXL Symbol Enc", "IDAutomationYHC39L Symbol",
        "IDAutomationYHC39M Symbol", "IDAutomationYHC39S Symbol", "IDAutomationYHC39XL Symbol Enc",
        "IDAutomationYHC39XS Symbol", "IDAutomationYHC39XXL Symbol Enc", "IDAutomationYHI25L Symbol",
        "IDAutomationYHI25M Symbol", "IDAutomationYHI25S Symbol", "IDAutomationYHI25XL Symbol Enc",
        "IDAutomationYHI25XS Symbol", "IDAutomationYI25L Symbol", "IDAutomationYI25M Symbol",
        "IDAutomationYI25S Symbol", "IDAutomationYI25XL Symbol Enc", "IDAutomationYI25XS Symbol",
        "IDAutomationYXC39L Symbol", "IDAutomationYXC39M Symbol", "IDAutomationYXC39S Symbol",
        "IDAutomationYXC39XL Symbol Enc", "IDAutomationYXC39XS Symbol", "IDAutomationYXC39XXL Symbol Enc",
        "IDAutomationYXHC39L Symbol", "IDAutomationYXHC39M Symbol", "IDAutomationYXHC39S Symbol",
        "IDAutomationYXHC39XL Symbol Enc", "IDAutomationYXHC39XS Symbol", "IDAutomationYXHC39XXL Symbol", "Impact",
        "Ink Free", "IrisUPC", "Iskoola Pota", "JasmineUPC", "Javanese Text", "KaiTi", "Kalinga", "Kartika",
        "Khmer UI", "KodchiangUPC", "Kokila", "Lao UI", "Latha", "Leelawadee", "Leelawadee UI",
        "Leelawadee UI Semilight", "Levenim MT", "LilyUPC", "Lucida Console", "Lucida Sans Unicode", "Malgun Gothic",
        "Malgun Gothic Semilight", "Mangal", "Marlett", "Meiryo", "Meiryo UI", "MICR", "Microsoft Himalaya",
        "Microsoft JhengHei", "Microsoft JhengHei Light", "Microsoft JhengHei UI", "Microsoft JhengHei UI Light",
        "Microsoft New Tai Lue", "Microsoft PhagsPa", "Microsoft Sans Serif", "Microsoft Tai Le", "Microsoft Uighur",
        "Microsoft YaHei", "Microsoft YaHei Light", "Microsoft YaHei UI", "Microsoft YaHei UI Light",
        "Microsoft Yi Baiti", "MingLiU", "MingLiU-ExtB", "MingLiU_HKSCS", "MingLiU_HKSCS-ExtB", "Miriam",
        "Miriam Fixed", "Mongolian Baiti", "MoolBoran", "MS Gothic", "MS Mincho", "MS PGothic", "MS PMincho",
        "MS UI Gothic", "MV Boli", "Myanmar Text", "Narkisim", "Neue Haas Grotesk Text Pro",
        "Neue Haas Grotesk Text Pro Blac", "Neue Haas Grotesk Text Pro Extr", "Neue Haas Grotesk Text Pro Ligh",
        "Neue Haas Grotesk Text Pro Medi", "Neue Haas Grotesk Text Pro Thin", "Neue Haas Grotesk Text Pro Ultr",
        "Nirmala UI", "Nirmala UI Semilight", "NSimSun", "Nyala", "OCRB", "Palatino Linotype",
        "Plantagenet Cherokee", "PMingLiU", "PMingLiU-ExtB", "Raavi", "Rockwell Nova", "Rockwell Nova Cond",
        "Rockwell Nova Cond Light", "Rockwell Nova Extra Bold", "Rockwell Nova Light", "Rod", "Sakkal Majalla",
        "Sanskrit Text", "Segoe Fluent Icons", "Segoe MDL2 Assets", "Segoe Print", "Segoe Script", "Segoe UI",
        "Segoe UI Black", "Segoe UI Emoji", "Segoe UI Historic", "Segoe UI Light", "Segoe UI Semibold",
        "Segoe UI Semilight", "Segoe UI Symbol", "Segoe UI Variable Display", "Segoe UI Variable Display Light",
        "Segoe UI Variable Display Semib", "Segoe UI Variable Display Semil", "Segoe UI Variable Small",
        "Segoe UI Variable Small Light", "Segoe UI Variable Small Semibol", "Segoe UI Variable Small Semilig",
        "Segoe UI Variable Text", "Segoe UI Variable Text Light", "Segoe UI Variable Text Semibold",
        "Segoe UI Variable Text Semiligh", "Shonar Bangla", "Shruti", "SimHei", "Simplified Arabic",
        "Simplified Arabic Fixed", "SimSun", "SimSun-ExtB", "SimSun-ExtG", "Sitka Banner", "Sitka Banner Semibold",
        "Sitka Display", "Sitka Display Semibold", "Sitka Heading", "Sitka Heading Semibold", "Sitka Small",
        "Sitka Small Semibold", "Sitka Subheading", "Sitka Subheading Semibold", "Sitka Text", "Sitka Text Semibold",
        "Sylfaen", "Symbol", "Tahoma", "Times New Roman", "Traditional Arabic", "Trebuchet MS", "Tunga",
        "UD Digi Kyokasho N-B", "UD Digi Kyokasho N-R", "UD Digi Kyokasho NK-B", "UD Digi Kyokasho NK-R",
        "UD Digi Kyokasho NP-B", "UD Digi Kyokasho NP-R", "Urdu Typesetting", "Utsaah", "Vani", "Verdana",
        "Verdana Pro", "Verdana Pro Black", "Verdana Pro Cond", "Verdana Pro Cond Black", "Verdana Pro Cond Light",
        "Verdana Pro Cond Semibold", "Verdana Pro Light", "Verdana Pro Semibold", "Vijaya", "Vrinda", "Webdings",
        "Wingdings", "Yu Gothic", "Yu Gothic Light", "Yu Gothic Medium", "Yu Gothic UI", "Yu Gothic UI Light",
        "Yu Gothic UI Semibold", "Yu Gothic UI Semilight", "Yu Mincho", "Yu Mincho Demibold", "Yu Mincho Light"
    ];

    private static readonly HashSet<string> Names =
        Listed.Select(TrimTtf).ToHashSet(StringComparer.OrdinalIgnoreCase);

    private static readonly string[] CutNames = Listed.Where(n => n.Length == LegacyFamilyNameLength).ToArray();

    /// <summary>True when Business Central online has <paramref name="fontName"/> installed (see the remarks).</summary>
    internal static bool IsInstalled(string fontName)
    {
        var name = TrimTtf(fontName.Trim());
        return Names.Contains(name)
            || CutNames.Any(cut => name.StartsWith(cut, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// The fonts embedded in <paramref name="main"/>'s font table (a <c>w:font</c> with an <c>w:embed…</c> font
    /// part), which Business Central uses whether or not it has them installed.
    /// </summary>
    internal static IReadOnlySet<string> EmbeddedIn(MainDocumentPart main) =>
        main.FontTablePart?.Fonts?.Elements<Font>()
            .Where(f => f.EmbedRegularFont is not null || f.EmbedBoldFont is not null
                || f.EmbedItalicFont is not null || f.EmbedBoldItalicFont is not null)
            .Select(f => f.Name?.Value)
            .OfType<string>()
            .ToHashSet(StringComparer.OrdinalIgnoreCase)
        ?? new HashSet<string>(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// The sentence a tool summary carries when <paramref name="fontName"/>, a font the caller asked for, will not
    /// print as named in Business Central online: neither installed there nor embedded in the layout. Empty
    /// otherwise.
    /// </summary>
    internal static string SummaryNote(string? fontName, MainDocumentPart? main)
    {
        if (fontName is null || IsInstalled(fontName) || (main is not null && EmbeddedIn(main).Contains(fontName)))
        {
            return string.Empty;
        }

        return $" '{fontName}' is not a font Business Central online has installed, and the layout does not embed "
            + "it, so BC prints that text in a substitute font (Microsoft's own layouts use Segoe UI). Pick an "
            + "installed font (learn.microsoft.com/dynamics365/business-central/ui-fonts), or embed this one from "
            + "Word.";
    }

    private static string TrimTtf(string name) =>
        name.EndsWith(".ttf", StringComparison.OrdinalIgnoreCase) ? name[..^4] : name;
}
