using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Domain;

/// <summary>
/// Builds the Word <c>PAGE</c>/<c>NUMPAGES</c> field-code run sequences every stock BC document header
/// uses for its "Page X / Y" chrome — the single most common piece of BC header content no tool could
/// emit before (<c>insert_text</c> is literal text only; GitHub issue #29). Like <see cref="SdtFactory"/>,
/// this type does no file I/O and inserts nothing — see <see cref="LayoutEditor.InsertPageNumber"/> for
/// placement.
/// </summary>
/// <remarks>
/// The exact shape mirrored here was extracted from the real header chrome of FOUR corpus layouts —
/// <c>StandardSalesQuote.docx</c> (header2/header3), <c>StandardPurchaseOrder.docx</c> (header2/header3),
/// <c>StandardSalesInvoiceVatSpec.docx</c> (header1) and <c>SalespersonCommission.docx</c> (header1) —
/// which all carry the identical five-run field construct per field, twice, joined by a literal
/// <c>" / "</c> run:
/// <code>
/// &lt;w:r&gt;&lt;w:fldChar w:fldCharType="begin"/&gt;&lt;/w:r&gt;
/// &lt;w:r&gt;&lt;w:instrText xml:space="preserve"&gt; PAGE  \* Arabic  \* MERGEFORMAT &lt;/w:instrText&gt;&lt;/w:r&gt;
/// &lt;w:r&gt;&lt;w:fldChar w:fldCharType="separate"/&gt;&lt;/w:r&gt;
/// &lt;w:r&gt;&lt;w:rPr&gt;&lt;w:noProof/&gt;&lt;/w:rPr&gt;&lt;w:t&gt;1&lt;/w:t&gt;&lt;/w:r&gt;
/// &lt;w:r&gt;&lt;w:fldChar w:fldCharType="end"/&gt;&lt;/w:r&gt;
/// </code>
/// then <c>&lt;w:r&gt;&lt;w:t xml:space="preserve"&gt; / &lt;/w:t&gt;&lt;/w:r&gt;</c>, then the same five
/// runs with <c> NUMPAGES  \* Arabic  \* MERGEFORMAT </c>. Three details are load-bearing, all four
/// captures agree on them, and each is asserted by a test:
/// <list type="bullet">
/// <item>the instruction text's EXACT spacing (leading/trailing space, TWO spaces before each <c>\*</c>),
/// preserved via <c>xml:space="preserve"</c> — Word tolerates variations, but this is the observed shape
/// and per ADR-0005 the observed shape is what gets emitted;</item>
/// <item>the cached result run (<c>1</c>) between <c>separate</c> and <c>end</c> carries
/// <c>w:noProof</c> — every capture has it, and it is what keeps the stale cached digit from being
/// spell-checked before Word recalculates the field;</item>
/// <item>the separator is a LITERAL run, not part of either field.</item>
/// </list>
/// The stock captures also carry decorative run properties on the field runs (<c>w:bCs</c> in 1304/1322,
/// an explicit Segoe UI font + size in 115) — per-layout styling residue of the paragraph each sits in,
/// omitted here exactly as <see cref="SdtFactory"/> omits its corpus examples' decorative <c>w:rPr</c>
/// (the caller's own <c>bold</c>/<c>fontSizePoints</c> knobs style the runs instead when asked).
/// <para>
/// The stock idiom's leading <c>Page_Lbl</c> label and the two literal spaces after it are deliberately
/// NOT part of this factory: they are ordinary <c>insert_label</c>/<c>insert_text</c> content the caller
/// composes (the label stays a translatable dataset binding that way), documented in the tool description.
/// </para>
/// </remarks>
public static class PageNumberFieldFactory
{
    /// <summary>The <c>PAGE</c> field instruction, exactly as all four corpus captures carry it.</summary>
    public const string PageInstruction = @" PAGE  \* Arabic  \* MERGEFORMAT ";

    /// <summary>The <c>NUMPAGES</c> field instruction, exactly as the corpus captures carry it.</summary>
    public const string NumPagesInstruction = @" NUMPAGES  \* Arabic  \* MERGEFORMAT ";

    /// <summary>The literal separator run between the two fields — <c>" / "</c> in every capture.</summary>
    public const string PageOfTotalSeparator = " / ";

    /// <summary>
    /// The runs for a bare current-page number: the five-run <c>PAGE</c> field construct (see the type
    /// remarks for the corpus evidence).
    /// </summary>
    public static IReadOnlyList<Run> BuildPageNumber() => BuildField(PageInstruction);

    /// <summary>
    /// The runs for the full stock "X / Y" shape: the <c>PAGE</c> field, the literal <c>" / "</c>
    /// separator run, and the <c>NUMPAGES</c> field — eleven runs total.
    /// </summary>
    public static IReadOnlyList<Run> BuildPageOfTotal()
    {
        var runs = new List<Run>(11);
        runs.AddRange(BuildField(PageInstruction));
        runs.Add(new Run(new Text(PageOfTotalSeparator) { Space = SpaceProcessingModeValues.Preserve }));
        runs.AddRange(BuildField(NumPagesInstruction));
        return runs;
    }

    /// <summary>
    /// One complete field construct: <c>begin</c>, the instruction (whitespace preserved verbatim — the
    /// spacing is part of the observed shape), <c>separate</c>, the cached result <c>1</c> under
    /// <c>w:noProof</c> (Word recalculates it on render; BC's own output does the same), and <c>end</c>.
    /// </summary>
    private static List<Run> BuildField(string instruction) =>
    [
        new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
        new Run(new FieldCode(instruction) { Space = SpaceProcessingModeValues.Preserve }),
        new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
        new Run(new RunProperties(new NoProof()), new Text("1")),
        new Run(new FieldChar { FieldCharType = FieldCharValues.End }),
    ];
}
