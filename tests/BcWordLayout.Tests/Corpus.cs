namespace BcWordLayout.Tests;

/// <summary>Locates the corpus .docx files copied next to the test assembly at build time.</summary>
public static class Corpus
{
    public const string SalesInvoice = "SalesInvoiceForSubscriptionBilling.docx";
    public const string InventoryOrderDetails = "InventoryOrderDetails.docx";
    public const string StandardStatement = "StandardStatement.docx";

    /// <summary><c>Standard_Sales_Quote/1304</c> — the layout the e2e scenario suite leans on most.</summary>
    public const string StandardSalesQuote = "StandardSalesQuote.docx";

    /// <summary><c>Standard_Purchase_Order/1322</c> — the purchase-side counterpart, with 6 header/footer parts.</summary>
    public const string StandardPurchaseOrder = "StandardPurchaseOrder.docx";

    // ---- 2026-08-01 base-app additions (captured from the Microsoft base app). Each carries a shape no
    // pre-existing corpus file has; the per-file docs below say what each one is here to cover.

    /// <summary>
    /// <c>Salesperson_Commission/115</c> — a TRUE base-app layout carrying the dedicated
    /// <c>&lt;Labels&gt;</c> data item whose label columns are suffixed <c>Label</c>/<c>Caption</c> rather than
    /// <c>Lbl</c>: the same shape <see cref="InventoryOrderDetails"/> covers, but without its custom-report
    /// provenance. Also one of the corpus's two <b>UTF-8</b>-encoded BC custom XML parts (the
    /// other is <see cref="SubcontractorDispatchList"/>; every other file is UTF-16 LE + BOM) and its only
    /// <c>w:tblHeader</c> repeat-header row on a repeater table.
    /// Validates with 0 errors and 1 warning: its dataset part has no itemProps (<c>single-bc-part</c>).
    /// </summary>
    public const string SalespersonCommission = "SalespersonCommission.docx";

    /// <summary>
    /// <c>Standard_Sales_Invoice/1306</c> (VAT-spec variant) — the corpus's ONLY <c>w:gridAfter</c> table,
    /// and it is the line-items repeater table itself (11 grid columns, 7 <c>gridSpan</c>, 3 <c>gridAfter</c>).
    /// This is the fixture long believed not to exist (see GitHub issue #9's history) — it is what LIFTED the blanket gridAfter
    /// refusal (see <c>TableStructureEditorTests</c>' gridAfter suite; <c>w:vMerge</c> remains the
    /// refused shape). Also carries two Microsoft-shipped defects: bindings whose <c>storeItemID</c>
    /// names a part that is not in the package, and a <c>w15:dataBinding</c> whose <c>w:prefixMappings</c>
    /// is a bare URI with no <c>xmlns:ns0=</c> declaration while its XPath still uses the <c>ns0:</c> prefix.
    /// </summary>
    public const string SalesInvoiceVatSpec = "StandardSalesInvoiceVatSpec.docx";

    /// <summary>
    /// <c>Job_Quote/1016</c> — the corpus's only MULTI-SECTION layout (4 <c>w:sectPr</c>, 6 header/footer
    /// parts), so the only one where "the first header part" and "this section's header part" differ.
    /// Also the only file mixing all three field-control shapes in one document: legacy
    /// <c>w:dataBinding</c>+<c>w:text</c>, <c>w15:dataBinding</c> with an alias/tag but no <c>w:text</c>,
    /// and <c>w15:dataBinding</c> with no alias/tag at all. Validates clean (0 errors, 0 warnings).
    /// </summary>
    public const string JobQuote = "JobQuote.docx";

    /// <summary>
    /// <c>Payment_Practice/685</c> — a base-app layout in which ALL 25 bindings are orphaned: 20 point at
    /// the report's OLD namespace (<c>Payment_Practice/590</c>) and all 25 name <c>storeItemID</c>s absent
    /// from the package, whose BC part has no <c>itemProps</c>/<c>DataStoreItem</c> at all. Kept because
    /// <c>validate_layout</c> used to report it as PASSING with zero findings. Sandbox-verified 2026-08-02
    /// (issue #1): BC REJECTS this capture at upload (<c>InvalidPrefixMapping</c> × 20) and accepts it once
    /// every binding is re-pointed to 685, so the 20 <c>binding-namespace</c> findings are Errors.
    /// </summary>
    public const string PaymentPracticeByPeriod = "PaymentPracticeByPeriod.docx";

    /// <summary>
    /// <c>Subcontractor_Dispatch_List/99000789</c> — the DEEPEST repeater nesting in the corpus: a
    /// straight 5-level chain (<c>Vendor</c> → <c>Work_Center</c> → <c>Prod_Order_Routing_Line</c> →
    /// <c>Prod_Order_Line</c> → <c>Prod_Order_Component</c>) with every level nested inside ONE 10-column
    /// body table. Structurally unlike <see cref="StandardStatement"/>'s 4-level shape, which branches
    /// across several tables and skips intermediate data items — this one is the pure-depth case, so it is
    /// the hardest test of the merge engine's per-level XPath re-anchoring. NOT a clean stock capture: the
    /// part claims 99000789 but 38 of its 43 bindings name 50000-range namespaces from a customized tenant
    /// (<c>Subcontractor_Dispatch_List/50030</c>, <c>FS_YSC_SubcontractorDispatch/50030</c>, even
    /// <c>Quantity_Explosion_of_BOM/50010</c>) — the same contamination signature that struck
    /// QuantityExplosionofBOM below. Kept for merge-depth coverage only; it fails <c>binding-namespace</c>
    /// with 38 Errors and must not be cited as evidence of how stock layouts ship (issue #1, 2026-08-02).
    /// </summary>
    public const string SubcontractorDispatchList = "SubcontractorDispatchList.docx";

    // ---- 2026-10-29 themable additions (source files from microsoft/BCApps, see PROVENANCE.md). Each colours
    // through the theme slot accent1 instead of literal colours.

    /// <summary>
    /// <c>Standard_Sales_Invoice/1306</c>, Blue Simple themable — the themable twin of
    /// <see cref="SalesInvoiceVatSpec"/>: <c>accent1</c> on 26 text colours, 8 cell fills and 25 table rules, and
    /// <c>Heading1</c>/<c>Heading2</c> title lines. Its embedded theme's <c>accent1</c> is <c>F24F4F</c>.
    /// </summary>
    public const string SalesInvoiceBlueSimpleThemable = "StandardSalesInvoiceBlueSimpleThemable.docx";

    /// <summary>
    /// <c>Standard_Sales_Shipment/1308</c>, Blue themable — carries one <c>w15:dataBinding</c> whose
    /// <c>w:prefixMappings</c> names a namespace with spaces (<c>…/Standard Sales - Shipment/1308/</c>, GitHub
    /// issue #154), and two captions its conversion left in a literal colour.
    /// </summary>
    public const string SalesShipmentBlueThemable = "StandardSalesShipmentBlueThemable.docx";

    /// <summary>
    /// <c>Standard_Purchase_Order/1322</c> themable — a colour-only twin of <see cref="StandardPurchaseOrder"/> (same
    /// styles and theme), and the corpus's only <c>w:fldSimple</c> page field.
    /// </summary>
    public const string StandardPurchaseOrderThemable = "StandardPurchaseOrderThemable.docx";

    // QuantityExplosionofBOM.docx was REMOVED from the corpus (2026-08-01). Business Central rejected it
    // on upload — InvalidPrefixMapping on nine bindings naming Quantity_Explosion_of_BOM/50000 and
    // QuantityExplosionofBOM/50013, neither of which is the report it claims to be (99000753) — and report
    // 99000753 does not run in the sandbox even with Microsoft's own layout. It was a corrupted/customized
    // capture, not a base-app one, so nothing may be concluded from it. Two coverage roles went with it and
    // are NOT replaced by another corpus file: it was the only layout with five customXml parts (the real
    // test of "find the BC part by namespace, never by part count/index") and the simpler of the two
    // w:gridAfter witnesses. Both want a SYNTHETIC fixture rather than another captured file.

    public static string Path(string fileName)
    {
        var full = System.IO.Path.Combine(AppContext.BaseDirectory, "corpus", fileName);
        if (!File.Exists(full))
        {
            throw new FileNotFoundException(
                $"Corpus file '{fileName}' not found at '{full}'. Ensure the .docx copy step ran.", full);
        }

        return full;
    }

    /// <summary>
    /// A temp copy of corpus layout <paramref name="fileName"/> without its external attached-template link: the state
    /// every tool save leaves a layout in, since Business Central refuses the link (see
    /// <see cref="BcWordLayout.Domain.AttachedTemplateLink"/>). For tests that expect a clean layout but edit through the
    /// Domain layer, which does not unlink, or read a stock layout before any tool has saved it. The caller deletes it.
    /// </summary>
    public static string UnlinkedCopy(string fileName)
    {
        var copy = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"bcwl-unlinked-{Guid.NewGuid():N}.docx");
        File.Copy(Path(fileName), copy, overwrite: true);
        using (var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(copy, true))
        {
            BcWordLayout.Domain.AttachedTemplateLink.Remove(doc);
        }

        return copy;
    }
}
