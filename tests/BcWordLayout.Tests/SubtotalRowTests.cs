using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;
using BcWordLayout.McpHost;
using BcWordLayout.McpHost.Tools;
using BcWordLayout.Merge;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Office2013Word = DocumentFormat.OpenXml.Office2013.Word;

namespace BcWordLayout.Tests;

/// <summary>
/// Covers GitHub issue #30: static per-group rows appended INSIDE an existing repeater's item —
/// <see cref="LayoutEditor.InsertSubtotalRow"/> and the <c>insert_subtotal_row</c> MCP tool. The target
/// shape is <c>SalespersonCommission.docx</c>'s (115) group structure: each salesperson's
/// <c>repeatingSectionItem</c> holds the group-header line row, the nested <c>Cust_Ledger_Entry</c>
/// detail repeater, an empty spacer row, and a bold subtotal row binding a sibling non-repeating
/// <c>Subtotals</c> item. The from-scratch build is exercised on a tool-built <c>/Header/Line</c>
/// repeater in the SalesInvoice corpus (which has both a real storeItemID and a nested
/// <c>AssemblyLine</c> child item), including the load-bearing merge semantics: the row must render once
/// per PARENT row, multiplying with groups but NOT with nested detail rows.
/// </summary>
[Collection("label-convention-seam")]
public class SubtotalRowTests
{
    private static string CopyOfCorpus(string corpusFile)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-subtotal-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    /// <summary>Builds the standard 3-column /Header/Line repeater table this suite appends group rows to.</summary>
    private static EditResult BuildLineRepeater(WordprocessingDocument doc) =>
        LayoutEditor.InsertRepeaterTable(
            doc,
            "/Header/Line",
            ["ItemNo_Line", "Description_Line", "Quantity_Line"],
            new Location { Type = LocationKind.DocumentEnd },
            new RepeaterTableOptions());

    // ---- the on-disk OOXML shape ----

    [Fact]
    public void Subtotal_row_is_appended_at_the_end_of_the_repeating_item()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            EditResult repeater;
            EditResult result;
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                repeater = BuildLineRepeater(doc);

                // The corpus group shape: a nested detail repeater first, then the trailing static rows.
                LayoutEditor.InsertRepeaterRow(doc, repeater.ControlId, "/Header/Line/AssemblyLine",
                    [new RepeaterRowCell { Span = 3, Columns = ["Description_AssemblyLine"] }]);

                result = LayoutEditor.InsertSubtotalRow(doc, repeater.ControlId,
                    [
                        new RepeaterRowCell { Span = 2, Columns = ["/Header/Line/Description_Line_Lbl"] },
                        new RepeaterRowCell { Span = 1, Columns = ["/Header/Line/AmountExcludingVAT_Line"], Alignment = "right" },
                    ],
                    new CellTextFormat { Bold = true });

                Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc));
                doc.MainDocumentPart!.Document!.Save();
            }

            Assert.Equal("insert_subtotal_row", result.Operation);
            Assert.Equal(0, result.ControlId);
            Assert.Equal("StaticRow", result.Kind);
            Assert.Equal(repeater.TableIndex, result.TableIndex);
            Assert.Contains("once per group", result.Summary, StringComparison.Ordinal);
            Assert.Contains("2 bound control(s)", result.Summary, StringComparison.Ordinal);

            using var reopened = WordprocessingDocument.Open(path, false);
            var body = reopened.MainDocumentPart!.Document!.Body!;
            var parent = body.Descendants<SdtRow>().Single(s => SdtInspector.ReadControlId(s) == repeater.ControlId);
            var itemContent = parent.GetFirstChild<SdtContentRow>()!.Elements<SdtRow>()
                .Single(s => s.SdtProperties?.GetFirstChild<Office2013Word.SdtRepeatedSectionItem>() is not null)
                .GetFirstChild<SdtContentRow>()!;

            // Item order is the corpus order: the line row, the nested detail repeater, then the static row.
            var children = itemContent.ChildElements.Where(e => e is TableRow or SdtRow).ToList();
            Assert.Equal(3, children.Count);
            Assert.IsType<TableRow>(children[0]);
            Assert.IsType<SdtRow>(children[1]);
            var subtotalRow = Assert.IsType<TableRow>(children[2]);

            // The row itself is INSIDE the repeatingSectionItem but is a bare w:tr (no wrapper of its own),
            // its label/field classification follows the path shape, and the amount cell is right-aligned
            // with the bold knob applied.
            var cells = subtotalRow.Elements<TableCell>().ToList();
            Assert.Equal(2, cells.Count);
            Assert.Equal(2, cells[0].GetFirstChild<TableCellProperties>()?.GetFirstChild<GridSpan>()?.Val?.Value);

            var caption = Assert.IsType<SdtRun>(cells[0].Descendants<SdtRun>().Single());
            Assert.Equal(ControlKind.Label, SdtInspector.ClassifyControlKind(caption));

            var amount = Assert.IsType<SdtRun>(cells[1].Descendants<SdtRun>().Single());
            Assert.Equal(ControlKind.Field, SdtInspector.ClassifyControlKind(amount));
            Assert.Equal(
                "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line[1]/ns0:AmountExcludingVAT_Line[1]",
                SdtInspector.ReadXPath(amount));
            Assert.Equal(
                JustificationValues.Right,
                cells[1].Descendants<Paragraph>().First().ParagraphProperties?.Justification?.Val?.Value);
            Assert.All(amount.Descendants<Run>(), r => Assert.NotNull(r.RunProperties?.Bold));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Repeated_calls_stack_rows_in_authoring_order_like_the_corpus_spacer_then_subtotal()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var repeater = BuildLineRepeater(doc);

                // The stock order: an all-spacer row first, then the subtotal row.
                LayoutEditor.InsertSubtotalRow(doc, repeater.ControlId, [new RepeaterRowCell { Span = 3 }]);
                LayoutEditor.InsertSubtotalRow(doc, repeater.ControlId,
                    [
                        new RepeaterRowCell { Span = 2 },
                        new RepeaterRowCell { Span = 1, Columns = ["/Header/Line/AmountExcludingVAT_Line"] },
                    ]);

                doc.MainDocumentPart!.Document!.Save();

                var parent = doc.MainDocumentPart.Document.Body!.Descendants<SdtRow>()
                    .Single(s => SdtInspector.ReadControlId(s) == repeater.ControlId);
                var itemContent = parent.GetFirstChild<SdtContentRow>()!.Elements<SdtRow>()
                    .Single(s => s.SdtProperties?.GetFirstChild<Office2013Word.SdtRepeatedSectionItem>() is not null)
                    .GetFirstChild<SdtContentRow>()!;

                var rows = itemContent.Elements<TableRow>().ToList();
                Assert.Equal(3, rows.Count); // line row, spacer, subtotal
                Assert.Empty(rows[1].Descendants<SdtRun>()); // the spacer binds nothing
                Assert.Single(rows[2].Descendants<SdtRun>());
            }
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- the load-bearing merge semantics: once per GROUP, not per detail row ----

    [Fact]
    public void Merged_preview_repeats_the_subtotal_once_per_parent_row_not_per_detail_row()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        var merged = Path.Combine(Path.GetTempPath(), $"bcwl-subtotal-merged-{Guid.NewGuid():N}.docx");
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var repeater = BuildLineRepeater(doc);
                LayoutEditor.InsertRepeaterRow(doc, repeater.ControlId, "/Header/Line/AssemblyLine",
                    [new RepeaterRowCell { Span = 3, Columns = ["Description_AssemblyLine"] }]);
                LayoutEditor.InsertSubtotalRow(doc, repeater.ControlId,
                    [
                        new RepeaterRowCell { Span = 2 },
                        new RepeaterRowCell { Span = 1, Columns = ["/Header/Line/AmountExcludingVAT_Line"] },
                    ]);
                doc.MainDocumentPart!.Document!.Save();
            }

            MergeEngine.Merge(path, merged);

            using var mergedDoc = WordprocessingDocument.Open(merged, false);
            var mergedBody = mergedDoc.MainDocumentPart!.Document!.Body!;

            // Counts must be scoped to the TOOL-BUILT table: the stock SalesInvoice body carries its own
            // /Header/Line repeater binding ItemNo_Line too. AmountExcludingVAT_Line is bound nowhere in
            // the stock layout, so it uniquely identifies the new table.
            var myTable = mergedBody.Descendants<SdtElement>()
                .First(s => SdtInspector.ReadXPath(s)?.Contains("ns0:AmountExcludingVAT_Line[", StringComparison.Ordinal) == true)
                .Ancestors<Table>().Last();

            static int CountByXPath(Table table, string fragment) => table.Descendants<SdtElement>()
                .Count(s => SdtInspector.ReadXPath(s)?.Contains(fragment, StringComparison.Ordinal) == true);

            // Sample data expands each repeater to N rows (default 3): the subtotal must appear exactly
            // once per PARENT Line row, while the nested AssemblyLine detail multiplied N-fold under each.
            // Fragments keep the trailing '[' so ItemNo_Line cannot also match the header's
            // ItemNo_Line_Lbl label control.
            var parentRows = CountByXPath(myTable, "ns0:ItemNo_Line[");
            var subtotals = CountByXPath(myTable, "ns0:AmountExcludingVAT_Line[");
            var detailRows = CountByXPath(myTable, "ns0:Description_AssemblyLine[");

            Assert.True(parentRows > 1, $"expected the /Header/Line rows to multiply, found {parentRows}");
            Assert.Equal(parentRows, subtotals);
            Assert.True(detailRows > subtotals,
                $"expected nested detail rows ({detailRows}) to outnumber subtotals ({subtotals}) - a "
                + "subtotal cloned per DETAIL row would make these equal or larger");
        }
        finally
        {
            File.Delete(path);
            if (File.Exists(merged))
            {
                File.Delete(merged);
            }
        }
    }

    // ---- the real corpus fixture: the stock shape stays readable and editable ----

    [Fact]
    public void A_spacer_row_can_be_appended_to_SalespersonCommission_the_issue_30_fixture()
    {
        // 115's own group shape IS the target evidence; its dataset part carries no storeItemID, so bound
        // cells cannot be authored into it - but the corpus's own spacer-row shape (an all-empty w:tr in
        // the item) binds nothing and must work.
        var path = CopyOfCorpus(Corpus.SalespersonCommission);
        try
        {
            int parentId;
            using (var doc = WordprocessingDocument.Open(path, false))
            {
                var inventory = LayoutReader.Read(doc);
                parentId = inventory.Controls
                    .Single(c => c.Kind == ControlKind.Repeater && c.ParentRepeater is null).SdtId!.Value;
            }

            EditResult result;
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                result = LayoutEditor.InsertSubtotalRow(doc, parentId, [new RepeaterRowCell { Span = 9 }]);
                Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc));
                doc.MainDocumentPart!.Document!.Save();
            }

            Assert.Contains("no bound cells (a spacer row)", result.Summary, StringComparison.Ordinal);

            using var reopened = WordprocessingDocument.Open(path, false);
            var parent = reopened.MainDocumentPart!.Document!.Body!.Descendants<SdtRow>()
                .Single(s => SdtInspector.ReadControlId(s) == parentId);
            var itemContent = parent.GetFirstChild<SdtContentRow>()!.Elements<SdtRow>()
                .Single(s => s.SdtProperties?.GetFirstChild<Office2013Word.SdtRepeatedSectionItem>() is not null)
                .GetFirstChild<SdtContentRow>()!;

            // The stock item already ends [_, ..., spacer, subtotal]; the new spacer is now last.
            var last = Assert.IsType<TableRow>(itemContent.ChildElements.Last(e => e is TableRow or SdtRow));
            var lastCell = Assert.Single(last.Elements<TableCell>());
            Assert.Equal(9, lastCell.GetFirstChild<TableCellProperties>()?.GetFirstChild<GridSpan>()?.Val?.Value);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- refusals ----

    [Fact]
    public void A_non_repeater_control_id_is_refused()
    {
        // An inline field's id is not a row-level repeater: it must be reported as not-found with the
        // "row-level repeater" wording — the same contract insert_repeater_row has for its parent id.
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            int fieldId;
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                fieldId = LayoutEditor.InsertField(
                    doc, "/Header/CustomerAddress1", new Location { Type = LocationKind.DocumentEnd }).ControlId;

                var ex = Assert.Throws<NotFoundException>(() =>
                    LayoutEditor.InsertSubtotalRow(doc, fieldId, [new RepeaterRowCell { Span = 1 }]));
                Assert.Contains("No row-level repeater control", ex.Message, StringComparison.Ordinal);
            }
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Spans_not_covering_the_parent_grid_are_refused()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var repeater = BuildLineRepeater(doc);

            var ex = Assert.Throws<ArgumentException>(() =>
                LayoutEditor.InsertSubtotalRow(doc, repeater.ControlId, [new RepeaterRowCell { Span = 2 }]));
            Assert.Contains("spans sum to 2", ex.Message, StringComparison.Ordinal);
            Assert.Contains("a subtotal row must cover the grid exactly", ex.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_row_level_alignment_is_refused_up_front()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var repeater = BuildLineRepeater(doc);

            var ex = Assert.Throws<ArgumentException>(() => LayoutEditor.InsertSubtotalRow(
                doc, repeater.ControlId, [new RepeaterRowCell { Span = 3 }],
                new CellTextFormat { Alignment = "right" }));
            Assert.Contains("per-cell alignments", ex.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void An_unknown_parent_id_is_reported_as_not_found()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            Assert.Throws<NotFoundException>(() =>
                LayoutEditor.InsertSubtotalRow(doc, 123456789, [new RepeaterRowCell { Span = 3 }]));
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- the insert_subtotal_row MCP tool ----

    [Fact]
    public void InsertSubtotalRow_tool_builds_the_corpus_group_shape_end_to_end()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var baseline = Assert.IsType<ValidationResultDto>(ReadTools.ValidateLayout(path, "quick").Data);

            var table = TableTools.InsertRepeaterTable(
                path, "/Header/Line", "ItemNo_Line,Description_Line,Quantity_Line", "documentEnd");
            Assert.True(table.Ok, table.Error?.Message);
            var tableDto = Assert.IsType<EditResultDto>(table.Data);

            // Spacer row, then the bold subtotal row - the corpus order, as the tool description says.
            var spacer = TableTools.InsertSubtotalRow(path, tableDto.ControlId, "3:-");
            Assert.True(spacer.Ok, spacer.Error?.Message);

            var subtotal = TableTools.InsertSubtotalRow(
                path, tableDto.ControlId,
                cells: "2:/Header/Line/Description_Line_Lbl,/Header/Line/AmountExcludingVAT_Line",
                alignments: "-,right",
                bold: true);
            Assert.True(subtotal.Ok, subtotal.Error?.Message);
            var dto = Assert.IsType<EditResultDto>(subtotal.Data);
            Assert.Equal("insert_subtotal_row", dto.Operation);
            Assert.Equal(0, dto.ControlId);
            Assert.Equal(baseline.ErrorCount, dto.QuickValidation.ErrorCount);

            // The deep (dry-run merge) level must not raise a single finding about the new bindings.
            var full = ReadTools.ValidateLayout(path, "full");
            Assert.True(full.Ok, full.Error?.Message);
            var fullDto = Assert.IsType<ValidationResultDto>(full.Data);
            Assert.DoesNotContain(fullDto.Findings, f =>
                (f.Message + f.Location).Contains("AmountExcludingVAT_Line", StringComparison.Ordinal));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertSubtotalRow_tool_reports_not_found_for_an_unknown_parent_id()
    {
        var response = TableTools.InsertSubtotalRow(Corpus.Path(Corpus.SalesInvoice), 123456789, "3:-");

        Assert.False(response.Ok);
        Assert.Equal("not_found", response.Error!.Code);
        Assert.False(string.IsNullOrWhiteSpace(response.Error.Hint));
    }
}
