using System.Xml.Linq;
using BcWordLayout.Domain;
using BcWordLayout.Merge;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

/// <summary>
/// The mock merge re-anchors a binding through EVERY repeater row enclosing it, not only the innermost one
/// (GitHub issue #104). It used to try the innermost row alone and read anything else from the document
/// root with every step indexed <c>[1]</c>. A field in a nested row bound to an OUTER row's column therefore
/// showed the first outer row's value in every group. A legitimately absolute binding in a row, such as a
/// <c>/Labels</c> caption, raised one <c>xpath-fallback</c> warning per generated row. What the merge still
/// cannot vouch for is reported once per binding, with a count: a <c>//</c> binding, and a field in another
/// repeater's rows that shows the first of several records of its own repeating item.
/// </summary>
public class AncestorRowReanchoringTests
{
    private const string Ns = "urn:microsoft-dynamics-nav/reports/TestReport/50000/";
    private const string Root = "/ns0:NavWordReportXmlPart[1]";
    private const string LineRepeater = Root + "/ns0:Header[1]/ns0:Line";
    private const string DetailRepeater = Root + "/ns0:Header[1]/ns0:Line[1]/ns0:Detail";
    private const string TagRepeater = Root + "/ns0:Header[1]/ns0:Line[1]/ns0:Tag";
    private const string DetailValue = Root + "/ns0:Header[1]/ns0:Line[1]/ns0:Detail[1]/ns0:Value[1]";
    private const string LineItemNo = Root + "/ns0:Header[1]/ns0:Line[1]/ns0:ItemNo[1]";
    private const string TagName = Root + "/ns0:Header[1]/ns0:Line[1]/ns0:Tag[1]/ns0:Name[1]";
    private const string HeaderDocumentNo = Root + "/ns0:Header[1]/ns0:DocumentNo[1]";
    private const string AnyValue = "//ns0:Value";
    private const string AnyNote = "//ns0:Note";

    /// <summary>
    /// InventoryOrderDetails' <c>/Item/Sales_Line/SalesHeaderBilltoName</c> control, a cell-level control in
    /// the nested Sales_Line row. The issue put <c>/Item/No_Item</c> into that row with <c>insert_field
    /// afterControl</c> on this control, which <c>insert_field</c> now refuses (nothing can follow a
    /// cell-level control inside its cell), so the test rebinds this control instead. Same shape: a field
    /// in the Sales_Line row bound to the enclosing Item row's column.
    /// </summary>
    private const int SalesHeaderBilltoNameControlId = 1956913051;

    private const string ItemNoXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Item[1]/ns0:No_Item[1]";

    private static string TempDocx(string tag) =>
        Path.Combine(Path.GetTempPath(), $"bcwl-reanchor-{tag}-{Guid.NewGuid():N}.docx");

    private static string CopyOfCorpus(string corpusFile)
    {
        var path = TempDocx("corpus");
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    private static string TextOf(SdtElement sdt) => string.Concat(sdt.Descendants<Text>().Select(t => t.Text));

    /// <summary>Every sdt under <paramref name="scope"/> whose binding is exactly <paramref name="xpath"/>, in
    /// document order. The merge never rewrites a binding, so clones keep the template's xpath.</summary>
    private static List<SdtElement> BoundTo(OpenXmlElement scope, string xpath) =>
        scope.Descendants<SdtElement>().Where(s => SdtInspector.ReadXPath(s) == xpath).ToList();

    /// <summary>Every sdt under <paramref name="scope"/> whose binding names exactly <paramref name="segments"/>,
    /// whatever its prefixes and predicates.</summary>
    private static List<SdtElement> BoundToPath(OpenXmlElement scope, params string[] segments) =>
        scope.Descendants<SdtElement>()
            .Where(s => BindingXPath.Segments(SdtInspector.ReadXPath(s)).SequenceEqual(segments))
            .ToList();

    private static string Repeater(string xpath, string rowContent, int id) =>
        "<w:sdt><w:sdtPr>"
        + $"<w:id w:val=\"{id}\"/>"
        + $"<w15:dataBinding w:xpath=\"{xpath}\" w:storeItemID=\"{SyntheticLayout.GoodItemId}\"/>"
        + "<w15:repeatingSection/></w:sdtPr><w:sdtContent>"
        + $"<w:sdt><w:sdtPr><w:id w:val=\"{id + 1}\"/><w15:repeatingSectionItem/></w:sdtPr>"
        + $"<w:sdtContent>{rowContent}</w:sdtContent></w:sdt>"
        + "</w:sdtContent></w:sdt>";

    private static string Field(string xpath) => SyntheticLayout.BoundField(xpath, SyntheticLayout.GoodItemId);

    /// <summary>
    /// A Line repeater whose row holds a <c>//</c> repeater and a nested Detail repeater. Each Detail row holds
    /// its own Value, the enclosing Line row's ItemNo, the top-level Header's DocumentNo, a <c>//</c> field,
    /// and a repeater over the enclosing Line row's Tag items.
    /// </summary>
    private static string ThreeLevelBody() =>
        Repeater(
            LineRepeater,
            Repeater(AnyNote, "<w:p><w:r><w:t>note</w:t></w:r></w:p>", id: 500)
            + Repeater(
                DetailRepeater,
                Field(DetailValue)
                + Field(LineItemNo)
                + Field(HeaderDocumentNo)
                + Field(AnyValue)
                + Repeater(TagRepeater, Field(TagName), id: 520),
                id: 510),
            id: 490);

    /// <summary>Two Lines, two Details each; Line A has one Tag, Line B two.</summary>
    private static SampleDataset ThreeLevelData()
    {
        XNamespace ns = Ns;
        XElement Line(string itemNo, string[] values, string[] tags) =>
            new(
                ns + "Line",
                new XElement(ns + "ItemNo", itemNo),
                values.Select(v => new XElement(ns + "Detail", new XElement(ns + "Value", v))),
                tags.Select(t => new XElement(ns + "Tag", new XElement(ns + "Name", t))));

        var xml = new XDocument(
            new XElement(
                ns + "NavWordReportXmlPart",
                new XElement(
                    ns + "Header",
                    new XElement(ns + "DocumentNo", "INV-1"),
                    Line("ITEM-A", ["A1", "A2"], ["ta"]),
                    Line("ITEM-B", ["B1", "B2"], ["tb1", "tb2"]),
                    new XElement(ns + "Note", new XElement(ns + "Text", "n1")))));

        return new SampleDataset { Xml = xml, Namespace = Ns };
    }

    [Fact]
    public void A_field_in_InventoryOrderDetails_nested_Sales_Line_row_bound_to_its_Item_row_shows_each_groups_own_item()
    {
        // The issue's reproduction: /Item/No_Item inside the nested Sales_Line row. Every group's lines used
        // to show item 1's number, with 9 xpath-fallback warnings (3 items x 3 lines).
        var path = CopyOfCorpus(Corpus.InventoryOrderDetails);
        var merged = TempDocx("iod-merged");
        try
        {
            using (var layout = WordprocessingDocument.Open(path, true))
            {
                var control = Assert.Single(
                    layout.MainDocumentPart!.Document!.Body!.Descendants<SdtElement>(),
                    s => SdtInspector.ReadControlId(s) == SalesHeaderBilltoNameControlId);
                SdtInspector.FindBinding(control)!.SetAttribute(new OpenXmlAttribute("w", "xpath", OoxmlNames.W, ItemNoXPath));
            }

            Assert.Equal(0, LayoutValidator.Quick(path).ErrorCount);

            var result = MergeEngine.Merge(path, merged, new MergeOptions { Rows = 3 });
            Assert.Equal(0, result.Stats.Unresolved);
            Assert.DoesNotContain(result.Warnings, w => w.Kind == "xpath-fallback");

            using var doc = WordprocessingDocument.Open(merged, false);
            var body = doc.MainDocumentPart!.Document!.Body!;
            var itemRepeater = Assert.Single(
                body.Descendants<SdtElement>(),
                s => SdtInspector.IsRepeater(s)
                    && BindingXPath.Segments(SdtInspector.ReadXPath(s)).SequenceEqual(new[] { "NavWordReportXmlPart", "Item" }));
            var itemRows = itemRepeater.Descendants<SdtElement>()
                .Where(s => SdtInspector.IsRepeaterItem(s) && SdtInspector.NearestRepeaterAncestor(s) == itemRepeater)
                .ToList();
            Assert.Equal(3, itemRows.Count);

            var groupItemNos = new List<string>();
            foreach (var itemRow in itemRows)
            {
                var noItemFields = BoundToPath(itemRow, "NavWordReportXmlPart", "Item", "No_Item");

                // The Item row's own header field sits directly in the Item row; the rebound control's copies
                // sit in the row's nested Sales_Line repeater, one per line.
                var header = Assert.Single(noItemFields, s => SdtInspector.NearestRepeaterAncestor(s) == itemRepeater);
                var inLines = noItemFields.Where(s => SdtInspector.NearestRepeaterAncestor(s) != itemRepeater).ToList();
                Assert.Equal(3, inLines.Count);

                var itemNo = TextOf(header);
                Assert.False(string.IsNullOrWhiteSpace(itemNo));
                Assert.All(inLines, s => Assert.Equal(itemNo, TextOf(s)));
                groupItemNos.Add(itemNo);
            }

            // Three different items, or "every line shows its own item" would hold vacuously.
            Assert.Equal(3, groupItemNos.Distinct(StringComparer.Ordinal).Count());
        }
        finally
        {
            File.Delete(path);
            File.Delete(merged);
        }
    }

    [Fact]
    public void SalespersonCommissions_Labels_caption_inside_the_Salesperson_row_validates_without_an_xpath_fallback()
    {
        // The stock layout's /Labels/CommissionPercentLabel sits inside the Salesperson row. Reading it from
        // the document root is exactly right - the caption lies outside every repeater - yet level=full used
        // to report it once per generated Salesperson row.
        var layout = CopyOfCorpus(Corpus.SalespersonCommission);
        var merged = TempDocx("spc-merged");
        try
        {
            var validation = FullValidator.Full(layout);
            Assert.DoesNotContain(validation.Findings, f => f.Check == "xpath-fallback");
            Assert.Equal(0, validation.ErrorCount);

            var result = MergeEngine.Merge(layout, merged, new MergeOptions { Rows = 3 });
            Assert.Empty(result.Warnings);

            using var doc = WordprocessingDocument.Open(merged, false);
            var captions = BoundToPath(
                doc.MainDocumentPart!.Document!, "NavWordReportXmlPart", "Labels", "CommissionPercentLabel");
            Assert.Equal(3, captions.Count);
            var caption = TextOf(captions[0]);
            Assert.False(string.IsNullOrWhiteSpace(caption));
            Assert.DoesNotContain("?»", caption, StringComparison.Ordinal);
            Assert.All(captions, s => Assert.Equal(caption, TextOf(s)));
        }
        finally
        {
            File.Delete(layout);
            File.Delete(merged);
        }
    }

    [Fact]
    public void Bindings_in_a_nested_row_re_anchor_to_the_deepest_enclosing_row_whose_path_they_extend()
    {
        var layoutPath = SyntheticLayout.Create(ThreeLevelBody());
        try
        {
            using var doc = WordprocessingDocument.Open(layoutPath, true);
            var result = MergeEngine.Merge(doc, ThreeLevelData());
            Assert.Equal(0, result.Stats.Unresolved);

            var body = doc.MainDocumentPart!.Document!.Body!;

            // Own row: each Detail row shows its own Value.
            Assert.Equal(new[] { "A1", "A2", "B1", "B2" }, BoundTo(body, DetailValue).Select(TextOf));

            // Enclosing Line row's column: each Detail row shows ITS Line's ItemNo. Re-anchoring against the
            // innermost row alone read /Header[1]/Line[1]/ItemNo[1] for all four - ITEM-A every time.
            Assert.Equal(new[] { "ITEM-A", "ITEM-A", "ITEM-B", "ITEM-B" }, BoundTo(body, LineItemNo).Select(TextOf));

            // A repeater over the enclosing Line row's Tag items expands THAT Line's tags in every Detail row:
            // Line A's one tag in each of its two Detail rows, then Line B's two tags in each of its two.
            Assert.Equal(new[] { "ta", "ta", "tb1", "tb2", "tb1", "tb2" }, BoundTo(body, TagName).Select(TextOf));

            // Outside every enclosing repeater: an absolute read, the same value everywhere.
            Assert.Equal(new[] { "INV-1", "INV-1", "INV-1", "INV-1" }, BoundTo(body, HeaderDocumentNo).Select(TextOf));
        }
        finally
        {
            File.Delete(layoutPath);
        }
    }

    [Fact]
    public void Only_bindings_that_cannot_be_re_anchored_warn_and_each_warns_once_with_its_evaluation_count()
    {
        var layoutPath = SyntheticLayout.Create(ThreeLevelBody());
        try
        {
            using var doc = WordprocessingDocument.Open(layoutPath, true);
            var result = MergeEngine.Merge(doc, ThreeLevelData());

            // The ancestor-row, own-row and outer-row-repeater bindings above raise nothing, nor does the
            // absolute DocumentNo (one Header, one DocumentNo: an unambiguous read); the two '//' bindings are
            // the only fallbacks here. The '//' field sits in each of the 4 Detail rows, the '//' repeater in
            // each of the 2 Line rows - one warning each, carrying that count, never one per row.
            var fallbacks = result.Warnings.Where(w => w.Kind == "xpath-fallback").ToList();
            Assert.Equal(2, fallbacks.Count);

            var repeater = Assert.Single(fallbacks, w => w.Location == $"document.xml: {AnyNote}");
            Assert.StartsWith("Repeater XPath", repeater.Message, StringComparison.Ordinal);
            Assert.Contains("(2 evaluation(s),", repeater.Message, StringComparison.Ordinal);

            var field = Assert.Single(fallbacks, w => w.Location == $"document.xml: {AnyValue}");
            Assert.StartsWith("Field XPath", field.Message, StringComparison.Ordinal);
            Assert.Contains("(4 evaluation(s),", field.Message, StringComparison.Ordinal);

            // Raised where the walk first met it: the Line row's '//' repeater comes before the Detail rows.
            Assert.True(result.Warnings.ToList().IndexOf(repeater) < result.Warnings.ToList().IndexOf(field));

            // Still read from the document root: the first Value in the dataset.
            var body = doc.MainDocumentPart!.Document!.Body!;
            Assert.Equal(new[] { "A1", "A1", "A1", "A1" }, BoundTo(body, AnyValue).Select(TextOf));
            Assert.Equal(2, BoundTo(body, AnyNote).Count);
        }
        finally
        {
            File.Delete(layoutPath);
        }
    }

    // ---- a field in another repeater's rows that shows the first of several records ----
    //
    // A field of a repeating data item placed in ANOTHER repeater's rows - a Line field in a ReportTotalsLine
    // row, what insert_column did before #82 - repeats with rows it does not belong to and shows the first
    // Line every time, in BC as in the mock. No quick-validation check catches it, so that root read warns.
    // Outside every repeater row a root read is genuinely absolute and stays silent: stock layouts read first
    // records there on purpose. So does a caption column, which reads the same in every record.

    private const string LineAmount = Root + "/ns0:Header[1]/ns0:Line[1]/ns0:Amount[1]";
    private const string LineDescription = Root + "/ns0:Header[1]/ns0:Line[1]/ns0:Description[1]";
    private const string LineItemNoLbl = Root + "/ns0:Header[1]/ns0:Line[1]/ns0:ItemNo_Lbl[1]";
    private const string LabelsCaption = Root + "/ns0:Labels[1]/ns0:TotalCaption[1]";
    private const string TotalsAmount = Root + "/ns0:Header[1]/ns0:Totals[1]/ns0:Amount[1]";
    private const string TotalsLineRepeater = Root + "/ns0:Header[1]/ns0:ReportTotalsLine";
    private const string TotalsLineDescription = Root + "/ns0:Header[1]/ns0:ReportTotalsLine[1]/ns0:Description[1]";

    private static string RowRepeater(string xpath, string cellContent, int id) =>
        "<w:sdt><w:sdtPr>"
        + $"<w:id w:val=\"{id}\"/>"
        + $"<w15:dataBinding w:xpath=\"{xpath}\" w:storeItemID=\"{SyntheticLayout.GoodItemId}\"/>"
        + "<w15:repeatingSection/></w:sdtPr><w:sdtContent>"
        + $"<w:sdt><w:sdtPr><w:id w:val=\"{id + 1}\"/><w15:repeatingSectionItem/></w:sdtPr>"
        + $"<w:sdtContent><w:tr><w:tc><w:tcPr/>{cellContent}</w:tc></w:tr></w:sdtContent></w:sdt>"
        + "</w:sdtContent></w:sdt>";

    /// <summary>
    /// A quote-style lines table: a heading row reading the first Line's caption column and its Description
    /// value; a Line repeater; and a ReportTotalsLine repeater whose row carries, besides its own column, a
    /// Line value field (the pre-#82 <c>insert_column</c> mistake), a Line caption column, an unambiguous
    /// <c>/Labels</c> caption and a single-record <c>/Header/Totals</c> field. Below the table, a Line field
    /// outside every repeater; in the page header, another.
    /// </summary>
    private static string LinesTableBody() =>
        "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"4000\"/></w:tblGrid>"
        + $"<w:tr><w:tc><w:tcPr/>{Field(LineItemNoLbl)}{Field(LineDescription)}</w:tc></w:tr>"
        + RowRepeater(LineRepeater, Field(LineItemNo), id: 600)
        + RowRepeater(
            TotalsLineRepeater,
            Field(TotalsLineDescription) + Field(LineAmount) + Field(LineItemNoLbl) + Field(LabelsCaption)
            + Field(TotalsAmount),
            id: 610)
        + "</w:tbl>"
        + Field(LineItemNo);

    private static SampleDataset LinesTableData(int lines)
    {
        XNamespace ns = Ns;
        var xml = new XDocument(
            new XElement(
                ns + "NavWordReportXmlPart",
                new XElement(ns + "Labels", new XElement(ns + "TotalCaption", "Total")),
                new XElement(
                    ns + "Header",
                    Enumerable.Range(1, lines).Select(i => new XElement(
                        ns + "Line",
                        new XElement(ns + "ItemNo", $"ITEM-{i}"),
                        new XElement(ns + "Description", $"Line {i}"),
                        new XElement(ns + "ItemNo_Lbl", "Item No."),
                        new XElement(ns + "Amount", $"{i}.00"))),
                    new XElement(ns + "ReportTotalsLine", new XElement(ns + "Description", "Subtotal")),
                    new XElement(ns + "ReportTotalsLine", new XElement(ns + "Description", "VAT")),
                    new XElement(ns + "Totals", new XElement(ns + "Amount", "99.00")))));

        return new SampleDataset { Xml = xml, Namespace = Ns };
    }

    [Fact]
    public void A_Line_field_in_a_ReportTotalsLine_row_warns_once_with_its_counts_while_reads_outside_every_row_stay_silent()
    {
        var layoutPath = SyntheticLayout.CreateWithHeader(LinesTableBody(), Field(LineItemNo));
        try
        {
            using var doc = WordprocessingDocument.Open(layoutPath, true);
            var result = MergeEngine.Merge(doc, LinesTableData(lines: 3));
            Assert.Equal(0, result.Stats.Unresolved);

            // The #82 shape is the one warning: the Line field repeats with each of the 2 totals lines,
            // showing Line 1 every time.
            var warning = Assert.Single(result.Warnings, w => w.Kind == "xpath-fallback");
            Assert.Equal($"document.xml: {LineAmount}", warning.Location);
            Assert.Equal(
                "Field '/Header/Line/Amount' lies outside every '/Header/Line' repeater, so it shows the first of "
                + "3 'Line' records (2 evaluation(s), one per generated copy of the control). Business Central "
                + "prints the first record there too; to show each record's own value, place the control in a "
                + $"'/Header/Line' repeater row: '{LineAmount}'.",
                warning.Message);

            // Silent, though each reads the first of 3 Lines: the heading row's caption AND value, the field
            // below the table and the page-header field all sit outside every repeater row (genuinely
            // absolute); inside the totals rows, the Line CAPTION reads the same in every record. /Labels and
            // /Header/Totals hold one record each, and the Line rows re-anchor.
            var body = doc.MainDocumentPart!.Document!.Body!;
            Assert.Equal(new[] { "1.00", "1.00" }, BoundTo(body, LineAmount).Select(TextOf));
            Assert.Equal(new[] { "Line 1" }, BoundTo(body, LineDescription).Select(TextOf));
            Assert.Equal(new[] { "Item No.", "Item No.", "Item No." }, BoundTo(body, LineItemNoLbl).Select(TextOf));
            Assert.Equal(new[] { "Total", "Total" }, BoundTo(body, LabelsCaption).Select(TextOf));
            Assert.Equal(new[] { "99.00", "99.00" }, BoundTo(body, TotalsAmount).Select(TextOf));
            Assert.Equal(new[] { "ITEM-1", "ITEM-2", "ITEM-3", "ITEM-1" }, BoundTo(body, LineItemNo).Select(TextOf));
            var header = doc.MainDocumentPart.HeaderParts.Single().Header!;
            Assert.Equal(new[] { "ITEM-1" }, BoundTo(header, LineItemNo).Select(TextOf));
        }
        finally
        {
            File.Delete(layoutPath);
        }
    }

    [Fact]
    public void The_same_layout_raises_nothing_when_the_data_holds_a_single_Line_record()
    {
        // A path through an item with ONE record is unambiguous wherever the control sits.
        var layoutPath = SyntheticLayout.CreateWithHeader(LinesTableBody(), Field(LineItemNo));
        try
        {
            using var doc = WordprocessingDocument.Open(layoutPath, true);
            var result = MergeEngine.Merge(doc, LinesTableData(lines: 1));

            Assert.Equal(0, result.Stats.Unresolved);
            Assert.DoesNotContain(result.Warnings, w => w.Kind == "xpath-fallback");
        }
        finally
        {
            File.Delete(layoutPath);
        }
    }
}
