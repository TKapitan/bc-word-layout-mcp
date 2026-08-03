using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace BcWordLayout.Tests;

/// <summary>
/// Builds minimal, deliberately malformed BC Word layout <c>.docx</c> files in memory (written to a
/// temp file) so the validator's negative paths can be exercised without shipping broken corpus docx.
/// The dataset custom XML part is a tiny <c>NavWordReportXmlPart</c> with a single <c>Header/CompanyName</c>
/// column; document.xml is assembled from raw sdt fragments the caller supplies.
/// </summary>
internal static class SyntheticLayout
{
    public const string GoodItemId = "{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}";
    public const string WrongItemId = "{11111111-2222-3333-4444-555555555555}";

    private const string Ns =
        "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
        + "xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\"";

    private const string DatasetXml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        + "<NavWordReportXmlPart xmlns=\"urn:microsoft-dynamics-nav/reports/TestReport/50000/\">"
        + "<BCReportInformation><CreationDateTime>2026-01-01</CreationDateTime></BCReportInformation>"
        + "<Header><CompanyName>Contoso</CompanyName></Header>"
        + "</NavWordReportXmlPart>";

    /// <summary>A bound field sdt (w:dataBinding) with the given xpath and storeItemID.</summary>
    public static string BoundField(string xpath, string storeItemId) =>
        "<w:sdt><w:sdtPr>"
        + $"<w:dataBinding w:xpath=\"{xpath}\" w:storeItemID=\"{storeItemId}\"/>"
        + "</w:sdtPr><w:sdtContent><w:p><w:r><w:t>x</w:t></w:r></w:p></w:sdtContent></w:sdt>";

    /// <summary>The dataset namespace <see cref="DatasetXml"/> declares — what a binding must name to be valid.</summary>
    public const string DatasetNamespace = "urn:microsoft-dynamics-nav/reports/TestReport/50000/";

    /// <summary>A different, plausible BC namespace — the shape of a binding left over from another report.</summary>
    public const string ForeignNamespace = "urn:microsoft-dynamics-nav/reports/OtherReport/50001/";

    /// <summary>
    /// A bound field sdt carrying an explicit <c>w:prefixMappings</c>, so
    /// <c>LayoutValidator</c>'s <c>binding-namespace</c> check has something to compare. <see cref="BoundField"/>
    /// deliberately emits no prefixMappings at all (the check must stay silent on those), so this is the
    /// variant to reach for when the namespace itself is what is under test.
    /// </summary>
    /// <param name="bareUri">
    /// When true the prefixMappings value is the RAW uri with no <c>xmlns:ns0='…'</c> declaration around it —
    /// a real base-app shape (StandardSalesInvoiceVatSpec.docx's Mini_Sales_Invoice bindings), not a synthetic
    /// curiosity, and the reason <c>SdtInspector.ExtractBcNamespace</c> matches by pattern rather than parsing
    /// an xmlns declaration.
    /// </param>
    public static string BoundFieldInNamespace(
        string xpath, string storeItemId, string namespaceUri, bool bareUri = false)
    {
        var mappings = bareUri ? namespaceUri : $"xmlns:ns0='{namespaceUri}' ";
        return "<w:sdt><w:sdtPr>"
            + $"<w:dataBinding w:prefixMappings=\"{mappings}\" w:xpath=\"{xpath}\" "
            + $"w:storeItemID=\"{storeItemId}\"/>"
            + "</w:sdtPr><w:sdtContent><w:p><w:r><w:t>x</w:t></w:r></w:p></w:sdtContent></w:sdt>";
    }

    /// <summary>A well-formed repeater (w15:repeatingSection) enclosing exactly one repeatingSectionItem.</summary>
    public static string ProperRepeater(string xpath, string storeItemId) =>
        "<w:sdt><w:sdtPr><w:id w:val=\"100\"/>"
        + $"<w15:dataBinding w:xpath=\"{xpath}\" w:storeItemID=\"{storeItemId}\"/>"
        + "<w15:repeatingSection/></w:sdtPr><w:sdtContent>"
        + "<w:sdt><w:sdtPr><w:id w:val=\"101\"/><w15:repeatingSectionItem/></w:sdtPr>"
        + "<w:sdtContent><w:p><w:r><w:t>row</w:t></w:r></w:p></w:sdtContent></w:sdt>"
        + "</w:sdtContent></w:sdt>";

    /// <summary>An orphaned repeatingSectionItem sdt with no enclosing repeatingSection.</summary>
    public static string OrphanRepeaterItem(int id = 900) =>
        $"<w:sdt><w:sdtPr><w:id w:val=\"{id}\"/><w15:repeatingSectionItem/></w:sdtPr>"
        + "<w:sdtContent><w:p><w:r><w:t>orphan</w:t></w:r></w:p></w:sdtContent></w:sdt>";

    /// <summary>A plain paragraph containing a single run of text — for <c>LocationResolver</c>'s AtText tests.</summary>
    public static string PlainParagraph(string text) =>
        $"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>";

    /// <summary>
    /// A plain paragraph containing a click-to-follow <c>w:hyperlink</c> run whose relationship id is
    /// <paramref name="relationshipId"/> — for <c>ExternalRelationshipStripper</c> tests proving a
    /// legitimate hyperlink survives the strip. Declares its own <c>xmlns:r</c> locally (this class's shared
    /// <see cref="Ns"/> constant only declares <c>w</c>/<c>w15</c>) so the fragment is self-contained; the
    /// caller adds the matching <c>HyperlinkRelationship</c> itself via the OpenXml SDK after
    /// <see cref="Create"/> writes the file (raw XML content alone cannot express a package relationship).
    /// </summary>
    public static string HyperlinkParagraph(string relationshipId, string text = "Click here") =>
        "<w:p><w:hyperlink xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
        + $"r:id=\"{relationshipId}\"><w:r><w:t>{text}</w:t></w:r></w:hyperlink></w:p>";

    /// <summary>
    /// An INLINE sdt (no binding, just a <c>w:id</c>) wrapped in its own paragraph — an <c>SdtRun</c>
    /// living inside a <c>w:p</c>, the same shape a real BC field control has. For
    /// <c>LocationResolver</c>'s <c>AfterControl</c> tests targeting an inline control.
    /// </summary>
    public static string InlineControlWithId(int id, string text = "x") =>
        "<w:p><w:sdt><w:sdtPr>"
        + $"<w:id w:val=\"{id}\"/>"
        + $"</w:sdtPr><w:sdtContent><w:r><w:t>{text}</w:t></w:r></w:sdtContent></w:sdt></w:p>";

    /// <summary>
    /// Same shape as <see cref="InlineControlWithId"/>, but the run is flanked by <c>w:proofErr</c>
    /// spell-check markers inside the sdt's own content — mirroring a REAL corpus shape (confirmed in
    /// SalesInvoiceForSubscriptionBilling.docx's header1.xml: a label control's content is
    /// <c>proofErr, run, proofErr</c>, not a bare run). For proving <c>LayoutEditor.RemoveControl</c>'s
    /// <c>keepText</c> unwrap moves every child of the content element, not just the one that looks like
    /// "the real content".
    /// </summary>
    public static string InlineControlWithProofErr(int id, string text = "x") =>
        "<w:p><w:sdt><w:sdtPr>"
        + $"<w:id w:val=\"{id}\"/>"
        + "</w:sdtPr><w:sdtContent><w:proofErr w:type=\"spellStart\"/>"
        + $"<w:r><w:t>{text}</w:t></w:r>"
        + "<w:proofErr w:type=\"spellEnd\"/></w:sdtContent></w:sdt></w:p>";

    /// <summary>
    /// A BLOCK-level sdt (no binding, just a <c>w:id</c>) whose own content is a paragraph — an
    /// <c>SdtBlock</c> sitting directly in the body's flow (a sibling of ordinary paragraphs), not
    /// wrapped inside one. For <c>LocationResolver</c>'s <c>AfterControl</c> tests targeting a
    /// block-level control.
    /// </summary>
    public static string BlockControlWithId(int id, string text = "x") =>
        "<w:sdt><w:sdtPr>"
        + $"<w:id w:val=\"{id}\"/>"
        + $"</w:sdtPr><w:sdtContent><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:sdtContent></w:sdt>";

    /// <summary>
    /// A minimal well-formed <c>w:tbl</c> with <paramref name="rows"/> rows and <paramref name="cols"/>
    /// columns; cell (r, c) contains one paragraph with text <c>"R{r}C{c}"</c>. For
    /// <c>LocationResolver</c>'s <c>TableCell</c> tests.
    /// </summary>
    public static string SimpleTable(int rows, int cols)
    {
        var grid = string.Concat(Enumerable.Repeat("<w:gridCol w:w=\"2000\"/>", cols));
        var sb = new System.Text.StringBuilder("<w:tbl><w:tblPr/><w:tblGrid>").Append(grid).Append("</w:tblGrid>");

        for (var r = 0; r < rows; r++)
        {
            sb.Append("<w:tr>");
            for (var c = 0; c < cols; c++)
            {
                sb.Append($"<w:tc><w:tcPr/><w:p><w:r><w:t>R{r}C{c}</w:t></w:r></w:p></w:tc>");
            }

            sb.Append("</w:tr>");
        }

        sb.Append("</w:tbl>");
        return sb.ToString();
    }

    /// <summary>
    /// A minimal well-formed 1×1 <c>w:tbl</c> whose <c>w:tblPr</c> carries a <c>w:tblStyle</c> reference to
    /// <paramref name="styleId"/> — for <c>LayoutValidator</c>'s <c>table-style-resolves</c> tests. Whether
    /// the reference actually resolves is the caller's business: pair it with
    /// <c>DefaultStylesScaffold.EnsureDefaultStyles</c> (which defines <c>TableGrid</c>) or leave the layout
    /// without a styles part entirely.
    /// </summary>
    public static string SimpleStyledTable(string styleId, string text = "styled") =>
        $"<w:tbl><w:tblPr><w:tblStyle w:val=\"{styleId}\"/></w:tblPr>"
        + "<w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid>"
        + $"<w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc></w:tr></w:tbl>";

    /// <summary>
    /// A 4-column <c>w:tbl</c> (1000 twips per grid column) whose three rows cover that grid three different
    /// ways — the <c>w:gridBefore</c>/<c>w:gridAfter</c> shape real base-app layouts use for a totals or
    /// filler row that stops short of the table's full width:
    /// <list type="bullet">
    /// <item>row 0 — four plain unit cells, no skipped columns (0 + 4 + 0).</item>
    /// <item>row 1 — <c>gridAfter=2</c>: two cells then two SKIPPED grid columns (0 + 2 + 2).</item>
    /// <item>row 2 — <c>gridBefore=1</c>: one skipped grid column then three cells (1 + 3 + 0).</item>
    /// </list>
    /// Every row still totals the grid's 4 columns, so <see cref="BcWordLayout.Domain.TableGridConsistencyGuard"/>
    /// sees a well-formed table — the point being that "well-formed" does NOT imply "every row starts at grid
    /// column 0", which is the assumption the column operations used to make.
    /// </summary>
    public static string TableWithSkippedGridCells()
    {
        var grid = string.Concat(Enumerable.Repeat("<w:gridCol w:w=\"1000\"/>", 4));

        static string Cell(string text, int width) =>
            $"<w:tc><w:tcPr><w:tcW w:w=\"{width}\" w:type=\"dxa\"/></w:tcPr>"
            + $"<w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>";

        return "<w:tbl><w:tblPr/><w:tblGrid>" + grid + "</w:tblGrid>"
            + "<w:tr>" + Cell("R0C0", 1000) + Cell("R0C1", 1000) + Cell("R0C2", 1000) + Cell("R0C3", 1000) + "</w:tr>"
            + "<w:tr><w:trPr><w:gridAfter w:val=\"2\"/></w:trPr>"
            + Cell("R1C0", 1000) + Cell("R1C1", 1000) + "</w:tr>"
            + "<w:tr><w:trPr><w:gridBefore w:val=\"1\"/></w:trPr>"
            + Cell("R2C0", 1000) + Cell("R2C1", 1000) + Cell("R2C2", 1000) + "</w:tr>"
            + "</w:tbl>";
    }

    /// <summary>
    /// A minimal well-formed <c>w:tbl</c> with one row: a CELL-LEVEL sdt (no binding, just a <c>w:id</c>)
    /// whose own content directly wraps a <c>w:tc</c> — an <c>SdtCell</c> sitting where an ordinary
    /// <c>w:tc</c> would (a sibling of the row's other cells, itself a child of <c>w:tr</c>) — plus one
    /// plain sibling cell. Mirrors the real corpus shape: BC field/label controls in a header table (e.g.
    /// <c>YourReference_Lbl</c>) are cell-level sdt, not inline or row-level. For
    /// <c>LocationResolver</c>'s <c>AfterControl</c> tests targeting a cell-level control.
    /// </summary>
    public static string TableWithCellLevelControl(int id, string text = "x") =>
        "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2000\"/><w:gridCol w:w=\"2000\"/></w:tblGrid>"
        + "<w:tr>"
        + "<w:sdt><w:sdtPr>"
        + $"<w:id w:val=\"{id}\"/>"
        + $"</w:sdtPr><w:sdtContent><w:tc><w:tcPr/><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc></w:sdtContent></w:sdt>"
        + "<w:tc><w:tcPr/><w:p><w:r><w:t>plain-cell</w:t></w:r></w:p></w:tc>"
        + "</w:tr></w:tbl>";

    /// <summary>
    /// A minimal well-formed <c>w:tbl</c> with one row of two cells; the FIRST cell's sole content is a
    /// BLOCK-level sdt (an <c>SdtBlock</c> whose content is one paragraph) sitting directly inside the
    /// <c>w:tc</c> — the exact shape BC uses for a line-items row's amount/quantity fields. Removing that
    /// control with keepText=false must NOT leave the cell empty (only its <c>w:tcPr</c>), which would be a
    /// silently-corrupt document. The second cell is a plain sibling cell.
    /// </summary>
    public static string TableWithBlockControlInCell(int id, string text = "x") =>
        "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2000\"/><w:gridCol w:w=\"2000\"/></w:tblGrid>"
        + "<w:tr>"
        + "<w:tc><w:tcPr/>"
        + "<w:sdt><w:sdtPr>"
        + $"<w:id w:val=\"{id}\"/>"
        + $"</w:sdtPr><w:sdtContent><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:sdtContent></w:sdt>"
        + "</w:tc>"
        + "<w:tc><w:tcPr/><w:p><w:r><w:t>plain-cell</w:t></w:r></w:p></w:tc>"
        + "</w:tr></w:tbl>";

    /// <summary>
    /// A minimal well-formed <c>w:tbl</c> containing a ROW-LEVEL sdt (no binding, just a <c>w:id</c>)
    /// whose own content directly wraps a <c>w:tr</c> — an <c>SdtRow</c> sitting where an ordinary
    /// <c>w:tr</c> would (a direct child of <c>w:tbl</c>). Mirrors real BC repeater controls
    /// (<c>w15:repeatingSection</c>/<c>w15:repeatingSectionItem</c>), which are also row-level sdt. For
    /// <c>LocationResolver</c>'s <c>AfterControl</c> tests proving a row-level target is rejected rather
    /// than producing invalid OOXML (a <c>w:tbl</c> cannot host a paragraph/block sdt as a direct sibling).
    /// </summary>
    public static string TableWithRowLevelControl(int id, string text = "x") =>
        "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid>"
        + "<w:sdt><w:sdtPr>"
        + $"<w:id w:val=\"{id}\"/>"
        + $"</w:sdtPr><w:sdtContent><w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc></w:tr></w:sdtContent></w:sdt>"
        + "</w:tbl>";

    /// <summary>
    /// A repeater whose single row template wraps one inner bound field sdt directly (sdt-in-sdt, no
    /// intervening paragraph — valid block content, and matches how real BC layouts nest a field inside a
    /// repeatingSectionItem). Lets merge-engine tests prove per-row re-anchoring of an inner binding.
    /// </summary>
    public static string RepeaterWithField(
        string repeaterXPath, string fieldXPath, string storeItemId, int repeaterId = 200, int itemId = 201) =>
        "<w:sdt><w:sdtPr>"
        + $"<w:id w:val=\"{repeaterId}\"/>"
        + $"<w15:dataBinding w:xpath=\"{repeaterXPath}\" w:storeItemID=\"{storeItemId}\"/>"
        + "<w15:repeatingSection/></w:sdtPr><w:sdtContent>"
        + "<w:sdt><w:sdtPr>"
        + $"<w:id w:val=\"{itemId}\"/>"
        + "<w15:repeatingSectionItem/></w:sdtPr>"
        + $"<w:sdtContent>{BoundField(fieldXPath, storeItemId)}</w:sdtContent></w:sdt>"
        + "</w:sdtContent></w:sdt>";

    /// <summary>
    /// A repeater TABLE — the row-level shape <c>SdtFactory.BuildRepeaterTable</c> emits and GitHub issue
    /// #19 targets: a single-column <c>w:tbl</c> whose first row (when <paramref name="headerRow"/> is
    /// true) is a static header <c>w:tr</c> marked <c>w:trPr/w:tblHeader</c>, followed by a row-level
    /// repeater — <c>SdtRow(w15:repeatingSection) &gt; SdtContentRow &gt; SdtRow(w15:repeatingSectionItem)
    /// &gt; SdtContentRow &gt; w:tr</c> — whose one data row's cell wraps a single bound field. For proving
    /// <c>MergeOptions.FlattenBindingsForRender</c> unwraps the row-level sdt shells so the cloned data
    /// rows end up plain <c>w:tr</c> SIBLINGS of the header row (Word fragments the table at surviving
    /// shells, losing <c>w:tblHeader</c> repetition). Pass <paramref name="headerRow"/> false for the
    /// degenerate table whose ONLY content is the repeater — with zero matched data rows the unwrap must
    /// then prune the rowless <c>w:tbl</c> itself.
    /// </summary>
    public static string RepeaterTable(
        string repeaterXPath, string fieldXPath, string storeItemId,
        bool headerRow = true, int repeaterId = 700, int itemId = 701) =>
        "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"4000\"/></w:tblGrid>"
        + (headerRow
            ? "<w:tr><w:trPr><w:tblHeader/></w:trPr>"
              + "<w:tc><w:tcPr/><w:p><w:r><w:t>No.</w:t></w:r></w:p></w:tc></w:tr>"
            : string.Empty)
        + "<w:sdt><w:sdtPr>"
        + $"<w:id w:val=\"{repeaterId}\"/>"
        + $"<w15:dataBinding w:xpath=\"{repeaterXPath}\" w:storeItemID=\"{storeItemId}\"/>"
        + "<w15:repeatingSection/></w:sdtPr><w:sdtContent>"
        + "<w:sdt><w:sdtPr>"
        + $"<w:id w:val=\"{itemId}\"/>"
        + "<w15:repeatingSectionItem/></w:sdtPr><w:sdtContent>"
        + $"<w:tr><w:tc><w:tcPr/>{BoundField(fieldXPath, storeItemId)}</w:tc></w:tr>"
        + "</w:sdtContent></w:sdt>"
        + "</w:sdtContent></w:sdt>"
        + "</w:tbl>";

    /// <summary>
    /// A two-level nested repeater: the outer row template (over <paramref name="outerXPath"/>) wraps an
    /// inner repeater (over <paramref name="innerXPath"/>, built via <see cref="RepeaterWithField"/>) whose
    /// own row wraps one bound field — for proving re-anchoring works across two nesting levels.
    /// </summary>
    public static string NestedRepeater(string outerXPath, string innerXPath, string fieldXPath, string storeItemId) =>
        "<w:sdt><w:sdtPr><w:id w:val=\"300\"/>"
        + $"<w15:dataBinding w:xpath=\"{outerXPath}\" w:storeItemID=\"{storeItemId}\"/>"
        + "<w15:repeatingSection/></w:sdtPr><w:sdtContent>"
        + "<w:sdt><w:sdtPr><w:id w:val=\"301\"/><w15:repeatingSectionItem/></w:sdtPr>"
        + $"<w:sdtContent>{RepeaterWithField(innerXPath, fieldXPath, storeItemId, repeaterId: 302, itemId: 303)}</w:sdtContent>"
        + "</w:sdt></w:sdtContent></w:sdt>";

    /// <summary>
    /// A repeater whose single row template wraps a PICTURE control (a <c>w:picture</c>-marked sdt with a
    /// minimal but schema-valid inline <c>w:drawing</c>/<c>wp:docPr</c>/blip — the real corpus shape, e.g.
    /// SalesInvoiceForSubscriptionBilling.docx's <c>PaymentServiceLogo</c>; the blip's <c>r:embed</c> need
    /// not resolve to a real relationship since <c>MergeEngine.ProcessPicture</c> unconditionally repoints
    /// it to a fresh placeholder <see cref="DocumentFormat.OpenXml.Packaging.ImagePart"/> it adds itself)
    /// followed by a <c>w:bookmarkStart</c>/<c>w:bookmarkEnd</c> pair — exactly the shape the clone-id fix
    /// targets: cloning this row N times used to leave every clone's <c>wp:docPr</c> id and bookmark id
    /// identical, tripping <c>OpenXmlValidator</c>'s <c>Sem_UniqueAttributeValue</c>.
    /// </summary>
    public static string RepeaterWithPictureAndBookmark(
        string repeaterXPath, string storeItemId, string bookmarkName = "TestBookmark",
        int repeaterId = 400, int itemId = 401, int pictureId = 402,
        uint docPrId = 1, string bookmarkId = "0") =>
        "<w:sdt><w:sdtPr>"
        + $"<w:id w:val=\"{repeaterId}\"/>"
        + $"<w15:dataBinding w:xpath=\"{repeaterXPath}\" w:storeItemID=\"{storeItemId}\"/>"
        + "<w15:repeatingSection/></w:sdtPr><w:sdtContent>"
        + "<w:sdt><w:sdtPr>"
        + $"<w:id w:val=\"{itemId}\"/>"
        + "<w15:repeatingSectionItem/></w:sdtPr><w:sdtContent>"
        + "<w:sdt><w:sdtPr>"
        + $"<w:id w:val=\"{pictureId}\"/>"
        + "<w:picture/></w:sdtPr><w:sdtContent>"
        + "<w:p><w:r><w:drawing>"
        + "<wp:inline"
        + " xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\""
        + " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""
        + " xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\""
        + " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""
        + " distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">"
        + "<wp:extent cx=\"100000\" cy=\"100000\"/>"
        + $"<wp:docPr id=\"{docPrId}\" name=\"Picture {docPrId}\"/>"
        + "<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect=\"1\"/></wp:cNvGraphicFramePr>"
        + "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"
        + "<pic:pic><pic:nvPicPr><pic:cNvPr id=\"0\" name=\"pic.png\"/><pic:cNvPicPr/></pic:nvPicPr>"
        + "<pic:blipFill><a:blip r:embed=\"rIdPlaceholder\"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>"
        + "<pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"100000\" cy=\"100000\"/></a:xfrm>"
        + "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr>"
        + "</pic:pic></a:graphicData></a:graphic></wp:inline>"
        + "</w:drawing></w:r></w:p>"
        + "</w:sdtContent></w:sdt>"
        + $"<w:p><w:bookmarkStart w:name=\"{bookmarkName}\" w:id=\"{bookmarkId}\"/>"
        + "<w:r><w:t>row</w:t></w:r>"
        + $"<w:bookmarkEnd w:id=\"{bookmarkId}\"/></w:p>"
        + "</w:sdtContent></w:sdt>"
        + "</w:sdtContent></w:sdt>";

    /// <summary>
    /// A plain paragraph containing ONLY a <c>w:bookmarkStart</c> (no matching end anywhere in this
    /// fragment) — for pairing with <see cref="RepeaterWithBookmarkEndOnly"/> to build a bookmark whose
    /// range SPANS a repeater boundary: this paragraph sits BEFORE the repeater in the body, while the
    /// matching <c>w:bookmarkEnd</c> lives inside the row template that gets cloned. Reproduces the
    /// reviewer-identified follow-up shape to the cloned-row id-uniqueness fix: a <c>BookmarkEnd</c> whose matching <c>BookmarkStart</c>
    /// is OUTSIDE the cloned row is never reached by a walk that only iterates the clone's own
    /// <c>BookmarkStart</c> elements, so without special handling every clone would repeat the SAME
    /// (unmatched) end id.
    /// </summary>
    public static string BookmarkStartOnly(string name, string id) =>
        $"<w:p><w:bookmarkStart w:name=\"{name}\" w:id=\"{id}\"/><w:r><w:t>before</w:t></w:r></w:p>";

    /// <summary>
    /// A repeater whose single row template wraps a <c>w:bookmarkEnd</c> with NO matching
    /// <c>w:bookmarkStart</c> inside the same row — see <see cref="BookmarkStartOnly"/>, whose companion
    /// start this end is meant to pair with (by <paramref name="bookmarkId"/>) OUTSIDE the repeater
    /// entirely.
    /// </summary>
    public static string RepeaterWithBookmarkEndOnly(
        string repeaterXPath, string storeItemId, string bookmarkId, int repeaterId = 600, int itemId = 601) =>
        "<w:sdt><w:sdtPr>"
        + $"<w:id w:val=\"{repeaterId}\"/>"
        + $"<w15:dataBinding w:xpath=\"{repeaterXPath}\" w:storeItemID=\"{storeItemId}\"/>"
        + "<w15:repeatingSection/></w:sdtPr><w:sdtContent>"
        + "<w:sdt><w:sdtPr>"
        + $"<w:id w:val=\"{itemId}\"/>"
        + "<w15:repeatingSectionItem/></w:sdtPr><w:sdtContent>"
        + $"<w:p><w:r><w:t>row</w:t></w:r><w:bookmarkEnd w:id=\"{bookmarkId}\"/></w:p>"
        + "</w:sdtContent></w:sdt>"
        + "</w:sdtContent></w:sdt>";

    /// <summary>
    /// Writes a temp .docx like <see cref="Create"/> (default Header/CompanyName dataset shape unless
    /// <paramref name="datasetXml"/> overrides it), but the BC dataset custom XML part has NO
    /// <see cref="CustomXmlPropertiesPart"/> at all — so <c>ds:itemID</c> is absent and
    /// <see cref="BcWordLayout.Domain.SchemaProvider.FromLayout(WordprocessingDocument)"/> resolves
    /// <c>StoreItemId</c> to null. Reproduces the one path through the real tool surface that reaches
    /// <c>SdtFactory</c>'s "schema.Report.StoreItemId is null/empty" guard: insert_field/insert_label/
    /// insert_repeater_table always build their schema via <c>SchemaProvider.FromLayout</c> against an
    /// already-open <c>.docx</c> (never <c>FromSchemaXml</c>, which has no <c>.docx</c>/properties part to
    /// begin with), so the only realistic trigger is a hand-crafted or externally-authored layout whose BC
    /// part lacks the properties part real BC-created layouts always carry.
    /// </summary>
    public static string CreateWithoutStoreItemId(string bodyFragments, string? datasetXml = null)
    {
        var documentXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<w:document {Ns}><w:body>{bodyFragments}"
            + "<w:sectPr/></w:body></w:document>";

        var path = System.IO.Path.Combine(
            System.IO.Path.GetTempPath(), $"bcwl-synth-noitemid-{Guid.NewGuid():N}.docx");

        using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            WriteRaw(main.GetStream(FileMode.Create, FileAccess.Write), documentXml);

            var cxp = main.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            WriteRaw(cxp.GetStream(FileMode.Create, FileAccess.Write), datasetXml ?? DatasetXml);

            // Deliberately no CustomXmlPropertiesPart/ds:itemID — see this method's own remarks.
        }

        return path;
    }

    /// <summary>
    /// Writes a temp .docx whose body contains the supplied sdt fragments and whose BC dataset part
    /// carries <paramref name="partItemId"/> as its ds:itemID. Pass <paramref name="datasetXml"/> to
    /// override the default (fixed Header/CompanyName-only) dataset shape — e.g. so
    /// <see cref="BcWordLayout.Domain.SchemaProvider"/> discovers a Header/Line or deeper shape matching
    /// the sdt fragments under test. Returns the temp file path.
    /// </summary>
    public static string Create(string bodyFragments, string partItemId = GoodItemId, string? datasetXml = null)
    {
        var documentXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<w:document {Ns}><w:body>{bodyFragments}"
            + "<w:sectPr/></w:body></w:document>";

        var itemPropsXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?>"
            + $"<ds:datastoreItem ds:itemID=\"{partItemId}\" "
            + "xmlns:ds=\"http://schemas.openxmlformats.org/officeDocument/2006/customXml\">"
            + "<ds:schemaRefs/></ds:datastoreItem>";

        var path = System.IO.Path.Combine(
            System.IO.Path.GetTempPath(), $"bcwl-synth-{Guid.NewGuid():N}.docx");

        using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            WriteRaw(main.GetStream(FileMode.Create, FileAccess.Write), documentXml);

            var cxp = main.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            WriteRaw(cxp.GetStream(FileMode.Create, FileAccess.Write), datasetXml ?? DatasetXml);

            var props = cxp.AddNewPart<CustomXmlPropertiesPart>();
            WriteRaw(props.GetStream(FileMode.Create, FileAccess.Write), itemPropsXml);
        }

        return path;
    }

    /// <summary>
    /// Writes a temp .docx like <see cref="Create"/>, but ALSO adds one header part containing
    /// <paramref name="headerFragments"/> raw sdt/paragraph fragments — for tests needing a control located
    /// specifically in a header (e.g. <see cref="BcWordLayout.Domain.LayoutValidator"/>'s
    /// repeater-in-header-footer check). The header part is added via a real <c>HeaderPart</c> relationship
    /// (so <c>MainDocumentPart.HeaderParts</c> enumerates it — all <see cref="BcWordLayout.Domain.LayoutReader"/>/
    /// <see cref="BcWordLayout.Domain.LayoutValidator"/> ever look at) without wiring a <c>w:headerReference</c>
    /// into the body's <c>w:sectPr</c> — unnecessary for these structural/binding checks, which never depend
    /// on section-level page layout (an orphaned-from-page-layout header part is still perfectly valid
    /// OOXML; <see cref="DocumentFormat.OpenXml.Validation.OpenXmlValidator"/> validates each part's own
    /// content model, not whether every part is reachable from a page's section properties).
    /// </summary>
    public static string CreateWithHeader(
        string bodyFragments, string headerFragments, string partItemId = GoodItemId, string? datasetXml = null)
    {
        var documentXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<w:document {Ns}><w:body>{bodyFragments}"
            + "<w:sectPr/></w:body></w:document>";

        var headerXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<w:hdr {Ns}>{headerFragments}</w:hdr>";

        var itemPropsXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?>"
            + $"<ds:datastoreItem ds:itemID=\"{partItemId}\" "
            + "xmlns:ds=\"http://schemas.openxmlformats.org/officeDocument/2006/customXml\">"
            + "<ds:schemaRefs/></ds:datastoreItem>";

        var path = System.IO.Path.Combine(
            System.IO.Path.GetTempPath(), $"bcwl-synth-header-{Guid.NewGuid():N}.docx");

        using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            WriteRaw(main.GetStream(FileMode.Create, FileAccess.Write), documentXml);

            var headerPart = main.AddNewPart<HeaderPart>();
            WriteRaw(headerPart.GetStream(FileMode.Create, FileAccess.Write), headerXml);

            var cxp = main.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            WriteRaw(cxp.GetStream(FileMode.Create, FileAccess.Write), datasetXml ?? DatasetXml);

            var props = cxp.AddNewPart<CustomXmlPropertiesPart>();
            WriteRaw(props.GetStream(FileMode.Create, FileAccess.Write), itemPropsXml);
        }

        return path;
    }

    private static void WriteRaw(Stream stream, string content)
    {
        using (stream)
        using (var writer = new StreamWriter(stream, new System.Text.UTF8Encoding(false)))
        {
            writer.Write(content);
        }
    }
}
