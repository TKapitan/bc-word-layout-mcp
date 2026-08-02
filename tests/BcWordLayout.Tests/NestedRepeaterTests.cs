using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;
using BcWordLayout.McpHost;
using BcWordLayout.McpHost.Tools;
using BcWordLayout.Merge;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

/// <summary>
/// Covers AUTHORING a nested repeater — a repeater table inside an outer repeater's data row.
/// Reading, validating, merging and previewing nested repeaters already worked 3–4 levels deep
/// (<see cref="XPathReanchor"/>); what was missing was a way to create one and, crucially, to know WHERE to
/// put it: <c>insert_repeater_table</c> reported only the new repeater control's id, so the caller had no
/// coordinate to address its data row with. It now reports the table's own index and data-row index too.
/// <para>
/// The corpus dataset used here is <c>SalesInvoiceForSubscriptionBilling.docx</c>'s real nesting:
/// <c>/Header/Line</c> with the child data item <c>/Header/Line/ShipmentLine</c>.
/// </para>
/// </summary>
public class NestedRepeaterTests
{
    private const string OuterItem = "/Header/Line";
    private const string InnerItem = "/Header/Line/ShipmentLine";

    private static string CopyOfCorpus(string corpusFile)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-nested-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    private static List<ValidationErrorInfo> OpenXmlErrors(WordprocessingDocument doc) =>
        new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc).ToList();

    /// <summary>True for a <c>w15:repeatingSectionItem</c> row — one expanded instance of a row template.</summary>
    private static bool IsRepeaterItem(SdtElement sdt) =>
        sdt.GetFirstChild<SdtProperties>()?.Elements()
            .Any(e => e.LocalName == "repeatingSectionItem" && e.NamespaceUri == OoxmlNames.W15) == true;

    [Fact]
    public void InsertRepeaterTable_reports_the_table_and_data_row_coordinates_a_nested_insert_needs()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = TableTools.InsertRepeaterTable(
                path, OuterItem, "ItemNo_Line,Description_Line", "documentEnd");
            Assert.True(response.Ok, response.Error?.Message);

            var dto = Assert.IsType<EditResultDto>(response.Data);
            Assert.NotNull(dto.TableIndex);
            Assert.NotNull(dto.DataRowIndex);
            Assert.Equal(1, dto.DataRowIndex); // header row 0, repeater data row 1

            // The reported coordinates really address the repeater's own data row.
            using var doc = WordprocessingDocument.Open(path, false);
            var table = TableGridNavigator.Tables(doc.MainDocumentPart!.Document!.Body!)[dto.TableIndex!.Value];
            var rows = TableGridNavigator.Rows(table);
            Assert.True(rows[dto.DataRowIndex!.Value].IsControlRow);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_repeater_table_authored_inside_another_repeaters_data_row_nests_and_validates()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var outer = Assert.IsType<EditResultDto>(TableTools.InsertRepeaterTable(
                path, OuterItem, "ItemNo_Line,Description_Line", "documentEnd").Data);

            // A column of its own for the nested table — the shape a real BC nested block has.
            var column = TableTools.InsertColumn(path, outer.TableIndex!.Value, mode: "plainText", headerText: "Shipments");
            Assert.True(column.Ok, column.Error?.Message);

            var inner = TableTools.InsertRepeaterTable(
                path, InnerItem, "PostingDate_ShipmentLine,Quantity_ShipmentLine", "tableCell",
                tableIndex: outer.TableIndex, row: outer.DataRowIndex, col: 2);
            Assert.True(inner.Ok, inner.Error?.Message);
            var innerDto = Assert.IsType<EditResultDto>(inner.Data);
            Assert.True(innerDto.QuickValidation.Passed, "a nested repeater must not introduce findings");

            using var reopened = WordprocessingDocument.Open(path, false);
            var inventory = LayoutReader.Read(reopened);

            // The inner repeater is reported as living inside the outer one.
            var innerControl = Assert.Single(inventory.Controls, c => c.SdtId == innerDto.ControlId);
            Assert.Equal(ControlKind.Repeater, innerControl.Kind);
            Assert.Equal(outer.ControlId, innerControl.ParentRepeater?.SdtId);
            Assert.Equal("/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line[1]/ns0:ShipmentLine", innerControl.XPath);

            Assert.Empty(TableGridConsistencyGuard.Find(reopened));
            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertRepeaterRow_adds_a_detail_row_inside_the_outer_repeaters_item()
    {
        // The STANDARD BC nesting shape (corpus-verified): per-line detail is a sibling ROW after the
        // line's own w:tr INSIDE the outer repeatingSectionItem, laid out on the SAME grid via gridSpans -
        // not a table hosted in a side column.
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var outer = Assert.IsType<EditResultDto>(TableTools.InsertRepeaterTable(
                path, OuterItem, "ItemNo_Line,Description_Line,LineAmount_Line", "documentEnd").Data);

            var response = TableTools.InsertRepeaterRow(
                path, outer.ControlId, InnerItem,
                cells: "-,2:PostingDate_ShipmentLine+Quantity_ShipmentLine",
                alignments: "-,left");
            Assert.True(response.Ok, response.Error?.Message);
            var dto = Assert.IsType<EditResultDto>(response.Data);
            Assert.Equal("Repeater", dto.Kind);
            Assert.Equal(outer.TableIndex, dto.TableIndex);
            Assert.True(dto.QuickValidation.Passed, "a detail row must not introduce findings");

            using var reopened = WordprocessingDocument.Open(path, false);
            var inventory = LayoutReader.Read(reopened);
            var inner = Assert.Single(inventory.Controls, c => c.SdtId == dto.ControlId);
            Assert.Equal(outer.ControlId, inner.ParentRepeater?.SdtId);
            Assert.Equal("/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line[1]/ns0:ShipmentLine", inner.XPath);

            // Structure: the outer ITEM's content is [line w:tr, nested repeater SdtRow] - the nested
            // repeater is a ROW SIBLING of the line row, and its own row covers the parent grid (1 + 2).
            var outerSdt = reopened.MainDocumentPart!.Document!.Body!.Descendants<SdtRow>()
                .Single(s => SdtInspector.ReadControlId(s) == outer.ControlId);
            var item = outerSdt.GetFirstChild<SdtContentRow>()!.Elements<SdtRow>().Single(IsRepeaterItem);
            var itemChildren = item.GetFirstChild<SdtContentRow>()!.ChildElements;
            Assert.IsType<TableRow>(itemChildren[0]);
            Assert.Equal(dto.ControlId, SdtInspector.ReadControlId((SdtRow)itemChildren[1]));

            var detailRow = itemChildren[1].Descendants<TableRow>().First();
            var detailCells = detailRow.Elements<TableCell>().ToList();
            Assert.Equal(2, detailCells.Count);
            Assert.Equal(2, detailCells[1].GetFirstChild<TableCellProperties>()!.GetFirstChild<GridSpan>()!.Val!.Value);
            // Two chained inline controls in the one content cell.
            Assert.Equal(2, detailCells[1].Descendants<SdtRun>().Count());

            Assert.Empty(TableGridConsistencyGuard.Find(reopened));
            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void An_authored_detail_row_expands_once_per_parent_row_in_a_merge()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        var mergedPath = Path.Combine(Path.GetTempPath(), $"bcwl-nested-merged-{Guid.NewGuid():N}.docx");
        try
        {
            var outer = Assert.IsType<EditResultDto>(TableTools.InsertRepeaterTable(
                path, OuterItem, "ItemNo_Line,Description_Line", "documentEnd").Data);
            Assert.True(TableTools.InsertRepeaterRow(
                path, outer.ControlId, InnerItem, cells: "-,PostingDate_ShipmentLine").Ok);

            var result = MergeEngine.Merge(path, mergedPath, new MergeOptions { Rows = 2 });
            Assert.Equal(0, result.Stats.Unresolved);

            using var merged = WordprocessingDocument.Open(mergedPath, false);
            var innerRepeaters = merged.MainDocumentPart!.Document!.Body!
                .Descendants<SdtRow>()
                .Where(s => SdtInspector.ReadXPath(s)?.EndsWith("ns0:ShipmentLine", StringComparison.Ordinal) == true)
                .ToList();
            Assert.Equal(2, innerRepeaters.Count); // one per expanded OUTER row
            Assert.Equal(4, innerRepeaters.Sum(r => r.Descendants<SdtRow>().Count(IsRepeaterItem))); // 2x2 detail rows
            Assert.Empty(OpenXmlErrors(merged));
        }
        finally
        {
            File.Delete(path);
            if (File.Exists(mergedPath))
            {
                File.Delete(mergedPath);
            }
        }
    }

    [Fact]
    public void InsertRepeaterRow_rejects_spans_that_do_not_cover_the_parent_grid()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var outer = Assert.IsType<EditResultDto>(TableTools.InsertRepeaterTable(
                path, OuterItem, "ItemNo_Line,Description_Line,LineAmount_Line", "documentEnd").Data);

            var response = TableTools.InsertRepeaterRow(
                path, outer.ControlId, InnerItem, cells: "-,PostingDate_ShipmentLine"); // covers 2 of 3
            Assert.False(response.Ok);
            Assert.Equal("invalid_argument", response.Error!.Code);
            Assert.Contains("grid", response.Error.Message);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertRepeaterRow_rejects_a_data_item_that_is_not_a_direct_child()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var outer = Assert.IsType<EditResultDto>(TableTools.InsertRepeaterTable(
                path, OuterItem, "ItemNo_Line,Description_Line", "documentEnd").Data);

            var response = TableTools.InsertRepeaterRow(
                path, outer.ControlId, "/Header/VATAmountLine", cells: "-,VATPct_VatAmountLine");
            Assert.False(response.Ok);
            Assert.Equal("invalid_argument", response.Error!.Code);
            Assert.Contains("DIRECT child", response.Error.Message);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void An_authored_nested_repeater_expands_row_by_row_in_a_merge()
    {
        // The proof that the authored bindings are the real thing: the merge engine re-anchors each inner
        // row to its own outer row's data node, so 2 outer rows x 2 inner rows = 4 inner instances.
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        var mergedPath = Path.Combine(Path.GetTempPath(), $"bcwl-nested-merged-{Guid.NewGuid():N}.docx");
        try
        {
            var outer = Assert.IsType<EditResultDto>(TableTools.InsertRepeaterTable(
                path, OuterItem, "ItemNo_Line,Description_Line", "documentEnd").Data);
            Assert.True(TableTools.InsertColumn(path, outer.TableIndex!.Value, mode: "plainText", headerText: "Shipments").Ok);
            var inner = Assert.IsType<EditResultDto>(TableTools.InsertRepeaterTable(
                path, InnerItem, "PostingDate_ShipmentLine,Quantity_ShipmentLine", "tableCell",
                tableIndex: outer.TableIndex, row: outer.DataRowIndex, col: 2).Data);

            var result = MergeEngine.Merge(path, mergedPath, new MergeOptions { Rows = 2 });
            Assert.Equal(0, result.Stats.Unresolved);

            using var merged = WordprocessingDocument.Open(mergedPath, false);
            var innerRepeaters = merged.MainDocumentPart!.Document!.Body!
                .Descendants<SdtRow>()
                .Where(s => SdtInspector.ReadXPath(s)?.EndsWith("ns0:ShipmentLine", StringComparison.Ordinal) == true)
                .ToList();

            // One inner repeater per expanded OUTER row...
            Assert.Equal(2, innerRepeaters.Count);

            // ...each expanded into its own rows: 2 x 2 = 4 inner data rows in total. Combined with
            // Unresolved == 0 above, that is the whole nesting contract — every inner row resolved against
            // its own outer row's data node rather than collapsing onto the first one.
            var innerDataRows = innerRepeaters.Sum(r => r.Descendants<SdtRow>().Count(IsRepeaterItem));
            Assert.Equal(4, innerDataRows);

            Assert.Empty(OpenXmlErrors(merged));
        }
        finally
        {
            File.Delete(path);
            if (File.Exists(mergedPath))
            {
                File.Delete(mergedPath);
            }
        }
    }
}
