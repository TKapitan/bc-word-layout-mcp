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
/// Covers <see cref="SdtFactory.BuildRepeaterTable"/> + <see cref="LayoutEditor.InsertRepeaterTable"/>
/// end-to-end against COPIES of the real <c>tests/corpus/SalesInvoiceForSubscriptionBilling.docx</c> corpus file's own
/// <c>/Header/Line</c> data item: OpenXML/BC validity, <see cref="LayoutReader"/> read-back shape, the
/// header label-resolution heuristic, unique id generation, and — the key proof — a
/// <see cref="MergeEngine"/> round-trip showing the newly inserted repeater expands and fills exactly like
/// a real BC-authored one.
/// </summary>
public class RepeaterTableTests
{
    // Real /Header/Line columns from tests/corpus/SalesInvoiceForSubscriptionBilling.docx (Standard_Sales_Invoice/1306):
    // ItemNo_Line / Description_Line each have a matching "<col>_Lbl" sibling in the same item;
    // TransHeaderAmount has no matching label anywhere in the whole schema (verified against the real
    // dataset custom XML part) - exercising both header-label-resolution branches in a single call.
    private const string LineItemPath = "/Header/Line";
    private static readonly string[] ThreeColumns = { "ItemNo_Line", "Description_Line", "TransHeaderAmount" };
    private const string RepeaterXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line";

    private static string CopyOfCorpus(string corpusFile)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-repeatertable-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    private static EditResult InsertThreeColumnTable(string path, RepeaterTableOptions? options = null)
    {
        using var doc = WordprocessingDocument.Open(path, true);
        var result = LayoutEditor.InsertRepeaterTable(
            doc, LineItemPath, ThreeColumns, new Location { Type = LocationKind.DocumentEnd },
            options ?? new RepeaterTableOptions());
        doc.MainDocumentPart!.Document!.Save();
        return result;
    }

    private static int? ReadId(SdtElement sdt) =>
        sdt.GetFirstChild<SdtProperties>()?.GetFirstChild<SdtId>()?.Val?.Value;

    private static List<int> AllSdtIds(OpenXmlElement root) =>
        root.Descendants<SdtElement>()
            .Select(ReadId)
            .Where(id => id.HasValue)
            .Select(id => id!.Value)
            .ToList();

    private static bool IsRepeaterItemSdt(SdtElement sdt) =>
        sdt.GetFirstChild<SdtProperties>()?.Elements()
            .Any(e => e.LocalName == "repeatingSectionItem" && e.NamespaceUri == OoxmlNames.W15) == true;

    /// <summary>
    /// Plain "zero OpenXmlValidator errors" assertion — kept as its own named helper (rather than an inline
    /// <c>Assert.Empty</c> at each call site) because both call sites below merge the SAME real corpus file
    /// (SalesInvoiceForSubscriptionBilling.docx), whose PaymentServiceLogo picture lives inside the
    /// pre-existing PaymentReportingArgument repeater and used to trip a duplicate <c>wp:docPr</c> id on
    /// every clone before the clone-id fix (<see cref="MergeEngine"/>'s row cloning now regenerates
    /// <c>wp:docPr</c> ids and bookmark <c>w:id</c>s per cloned row) — nothing to do with the NEW repeater
    /// each test itself inserts and checks.
    /// </summary>
    private static void AssertNoOpenXmlErrors(WordprocessingDocument doc)
    {
        var errors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc).ToList();

        Assert.True(errors.Count == 0,
            "expected zero validation errors; found: "
            + string.Join(" | ", errors.Select(e => $"{e.Path?.XPath}: {e.Description}")));
    }

    // ---- structural validity ----

    [Fact]
    public void InsertRepeaterTable_passes_OpenXmlValidator_and_LayoutValidator_Quick_with_zero_errors()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var result = InsertThreeColumnTable(path);

            Assert.Equal("insert_repeater_table", result.Operation);
            Assert.Equal("Repeater", result.Kind);
            Assert.Equal(3, result.ColumnCount);
            Assert.Equal(RepeaterXPath, result.XPath);
            Assert.Equal("#Nav: /Header/Line", result.Alias);
            Assert.Equal("document.xml", result.Part);

            using var reopened = WordprocessingDocument.Open(path, false);
            var openXmlErrors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened).ToList();
            Assert.Empty(openXmlErrors);

            var quick = LayoutValidator.Quick(reopened);
            Assert.True(quick.Passed, "errors: " + string.Join(" | ", quick.Findings
                .Where(f => f.Severity == FindingSeverity.Error).Select(f => f.Message)));
            Assert.Equal(0, quick.ErrorCount);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- B32: the BC-native look (no drawn grid; a per-cell rule under the header row) ----

    [Fact]
    public void InsertRepeaterTable_defaults_to_the_BC_look_no_table_grid_and_a_per_cell_header_rule()
    {
        // Corpus-verified: real BC lines tables carry NO w:tblBorders at all, and every w:tcBorders in the
        // whole corpus is one of two shapes - a single 1/2-pt bottom rule (under a header row) or the same
        // as a top rule (above a totals block). An explicit full border grid, which this tool used to draw
        // unconditionally, was the one visible fidelity gap left in the from-scratch authoring exercise.
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var result = InsertThreeColumnTable(path);

            using var reopened = WordprocessingDocument.Open(path, false);
            var table = FindInsertedTable(reopened, result.ControlId);

            Assert.Null(table.GetFirstChild<TableProperties>()!.GetFirstChild<TableBorders>());

            var headerCells = table.Elements<TableRow>().First().Elements<TableCell>().ToList();
            Assert.Equal(3, headerCells.Count);
            foreach (var cell in headerCells)
            {
                var borders = cell.TableCellProperties!.TableCellBorders!;
                var bottom = borders.BottomBorder!;
                Assert.Equal(BorderValues.Single, bottom.Val!.Value);
                Assert.Equal(4u, bottom.Size!.Value);
                Assert.Equal("auto", bottom.Color!.Value);

                // Only the header rule — the other three edges stay undrawn, like the corpus.
                Assert.Null(borders.TopBorder);
                Assert.Null(borders.LeftBorder);
                Assert.Null(borders.RightBorder);
            }

            // The data row is left completely undrawn.
            var dataCells = table.Descendants<TableRow>().Last().Elements<TableCell>().ToList();
            Assert.All(dataCells, c => Assert.Null(c.TableCellProperties?.TableCellBorders));

            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertRepeaterTable_look_grid_still_draws_the_full_explicit_border_grid()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var result = InsertThreeColumnTable(path, new RepeaterTableOptions { Look = TableBorderLook.Grid });

            using var reopened = WordprocessingDocument.Open(path, false);
            var table = FindInsertedTable(reopened, result.ControlId);

            var borders = table.GetFirstChild<TableProperties>()!.GetFirstChild<TableBorders>()!;
            Assert.Equal(BorderValues.Single, borders.TopBorder!.Val!.Value);
            Assert.Equal(BorderValues.Single, borders.InsideHorizontalBorder!.Val!.Value);
            Assert.Equal(BorderValues.Single, borders.InsideVerticalBorder!.Val!.Value);

            // The grid look draws everything at table level, so no per-cell header rule is added.
            var headerCells = table.Elements<TableRow>().First().Elements<TableCell>();
            Assert.All(headerCells, c => Assert.Null(c.TableCellProperties?.TableCellBorders));

            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    /// <summary>The <c>w:tbl</c> holding the repeater row with <paramref name="repeaterControlId"/>.</summary>
    private static Table FindInsertedTable(WordprocessingDocument doc, int repeaterControlId) =>
        doc.MainDocumentPart!.Document!.Body!.Descendants<SdtRow>()
            .First(s => ReadId(s) == repeaterControlId)
            .Ancestors<Table>()
            .First();

    [Fact]
    public void InsertRepeaterTable_reads_back_as_a_single_Repeater_with_the_unindexed_xpath_and_Field_cells()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var result = InsertThreeColumnTable(path);

            using var reopened = WordprocessingDocument.Open(path, false);
            var inventory = LayoutReader.Read(reopened);

            var repeater = Assert.Single(inventory.Controls, c => c.SdtId == result.ControlId);
            Assert.Equal(ControlKind.Repeater, repeater.Kind);
            Assert.Equal(RepeaterXPath, repeater.XPath);
            Assert.True(repeater.UsesW15Binding);

            var cellFields = inventory.Controls
                .Where(c => c.ParentRepeater is not null && c.ParentRepeater.SdtId == result.ControlId)
                .ToList();
            Assert.Equal(3, cellFields.Count);
            Assert.All(cellFields, c => Assert.Equal(ControlKind.Field, c.Kind));
            Assert.Contains(cellFields, c =>
                c.XPath == "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line[1]/ns0:ItemNo_Line[1]");
            Assert.Contains(cellFields, c =>
                c.XPath == "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line[1]/ns0:Description_Line[1]");
            Assert.Contains(cellFields, c =>
                c.XPath == "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line[1]/ns0:TransHeaderAmount[1]");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertRepeaterTable_header_row_resolves_labels_where_present_and_uses_static_text_otherwise()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var result = InsertThreeColumnTable(path);

            using var reopened = WordprocessingDocument.Open(path, false);
            var body = reopened.MainDocumentPart!.Document!.Body!;

            // Scoped to the NEWLY INSERTED table's own ids only - the corpus's OWN pre-existing Header/Line
            // table legitimately has real ItemNo_Line_Lbl/Description_Line_Lbl label controls of its own
            // (bound to the identical xpaths), which a whole-document Contains() check would satisfy
            // trivially regardless of whether OUR label-resolution heuristic did anything at all.
            var insertedTable = body.Elements<Table>().Last();
            var insertedIds = AllSdtIds(insertedTable).ToHashSet();

            var inventory = LayoutReader.Read(reopened);
            var ourControls = inventory.Controls.Where(c => c.SdtId.HasValue && insertedIds.Contains(c.SdtId.Value)).ToList();

            // ItemNo_Line / Description_Line each have a matching "_Lbl" sibling -> bound label controls.
            // These are NOT nested inside our repeater - real BC shape: the header w:tr is a sibling of the
            // repeater's SdtRow, both direct children of the same w:tbl (see BuildRepeaterTable's remarks).
            Assert.Contains(ourControls, c =>
                c.Kind == ControlKind.Label &&
                c.XPath == "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line[1]/ns0:ItemNo_Line_Lbl[1]" &&
                (c.ParentRepeater is null || c.ParentRepeater.SdtId != result.ControlId));
            Assert.Contains(ourControls, c =>
                c.Kind == ControlKind.Label &&
                c.XPath == "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line[1]/ns0:Description_Line_Lbl[1]");

            // Exactly 2 label controls in OUR table (not 3) - TransHeaderAmount has no matching label anywhere
            // in the schema, so it must have gotten no label control at all (static text fallback instead).
            Assert.Equal(2, ourControls.Count(c => c.Kind == ControlKind.Label));

            // The static header cell's own visible text is the humanized column name.
            Assert.Contains(insertedTable.Descendants<Text>(), t => t.Text == "Trans Header Amount");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertRepeaterTable_headerFromLabels_false_uses_static_text_for_every_column()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            InsertThreeColumnTable(path, new RepeaterTableOptions { HeaderFromLabels = false });

            using var reopened = WordprocessingDocument.Open(path, false);
            var body = reopened.MainDocumentPart!.Document!.Body!;

            // Scoped to the NEWLY INSERTED table only (last direct-child Table, documentEnd insertion) -
            // the corpus's OWN pre-existing Header/Line table legitimately has real ItemNo_Line_Lbl/
            // Description_Line_Lbl label controls of its own, which a whole-document check would wrongly
            // match against.
            var insertedTable = body.Elements<Table>().Last();
            Assert.DoesNotContain("Lbl", insertedTable.OuterXml, StringComparison.Ordinal);

            Assert.Contains(insertedTable.Descendants<Text>(), t => t.Text == "Item No Line");
            Assert.Contains(insertedTable.Descendants<Text>(), t => t.Text == "Description Line");
            Assert.Contains(insertedTable.Descendants<Text>(), t => t.Text == "Trans Header Amount");

            var openXmlErrors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened).ToList();
            Assert.Empty(openXmlErrors);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- id uniqueness ----

    [Fact]
    public void InsertRepeaterTable_generated_ids_are_mutually_unique_and_do_not_collide_with_pre_existing_ids()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            HashSet<int> preExistingIds;
            using (var probe = WordprocessingDocument.Open(path, false))
            {
                preExistingIds = AllSdtIds(probe.MainDocumentPart!.Document!.Body!).ToHashSet();
            }

            Assert.NotEmpty(preExistingIds); // sanity: the corpus does have pre-existing ids to collide with.

            InsertThreeColumnTable(path);

            using var reopened = WordprocessingDocument.Open(path, false);
            var body = reopened.MainDocumentPart!.Document!.Body!;

            // documentEnd appends as the last direct-child Table of the body, so this is exactly (and only)
            // the table InsertRepeaterTable just built.
            var insertedTable = body.Elements<Table>().Last();
            var newIds = AllSdtIds(insertedTable);

            // 2 header labels (ItemNo_Line_Lbl, Description_Line_Lbl) + 3 data fields + 1
            // repeatingSectionItem + 1 repeatingSection = 7 ids issued for this table.
            Assert.Equal(7, newIds.Count);
            Assert.Equal(newIds.Count, newIds.Distinct().Count());
            Assert.Empty(newIds.Intersect(preExistingIds));
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- column widths ----

    [Fact]
    public void InsertRepeaterTable_explicit_columnWidths_are_applied_to_tblGrid_and_cells()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                LayoutEditor.InsertRepeaterTable(
                    doc, LineItemPath, ThreeColumns, new Location { Type = LocationKind.DocumentEnd },
                    new RepeaterTableOptions { ColumnWidths = new[] { 1000, 2000, 3000 } });
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var body = reopened.MainDocumentPart!.Document!.Body!;
            var table = body.Elements<Table>().Last();

            var gridWidths = table.GetFirstChild<TableGrid>()!.Elements<GridColumn>()
                .Select(g => g.Width!.Value)
                .ToList();
            Assert.Equal(new[] { "1000", "2000", "3000" }, gridWidths);

            var openXmlErrors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened).ToList();
            Assert.Empty(openXmlErrors);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void BuildRepeaterTable_columnWidths_count_mismatch_throws_ArgumentException()
    {
        var schema = SchemaProvider.FromLayout(Corpus.Path(Corpus.SalesInvoice));

        var ex = Assert.Throws<ArgumentException>(() => SdtFactory.BuildRepeaterTable(
            schema, LineItemPath, ThreeColumns,
            new RepeaterTableOptions { ColumnWidths = new[] { 1000, 2000 } }, nextId: () => 1));

        Assert.Contains("ColumnWidths", ex.Message);
    }

    // ---- negative: dataItemPath validation ----

    [Fact]
    public void InsertRepeaterTable_nonexistent_dataItem_throws_ArgumentException_and_leaves_the_document_unmodified()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var body = doc.MainDocumentPart!.Document!.Body!;
            var tableCountBefore = body.Elements<Table>().Count();

            Assert.Throws<ArgumentException>(() => LayoutEditor.InsertRepeaterTable(
                doc, "/Header/ThisDataItemDoesNotExistAnywhere", ThreeColumns,
                new Location { Type = LocationKind.DocumentEnd }, new RepeaterTableOptions()));

            Assert.Equal(tableCountBefore, body.Elements<Table>().Count());
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void BuildRepeaterTable_dataItem_that_is_a_leaf_column_throws_ArgumentException()
    {
        var schema = SchemaProvider.FromLayout(Corpus.Path(Corpus.SalesInvoice));

        var ex = Assert.Throws<ArgumentException>(() => SdtFactory.BuildRepeaterTable(
            schema, "/Header/CustomerAddress1", ThreeColumns, new RepeaterTableOptions(), nextId: () => 1));

        Assert.Contains("leaf column", ex.Message);
    }

    [Fact]
    public void BuildRepeaterTable_dataItem_that_is_the_system_item_throws_ArgumentException()
    {
        var schema = SchemaProvider.FromLayout(Corpus.Path(Corpus.SalesInvoice));

        var ex = Assert.Throws<ArgumentException>(() => SdtFactory.BuildRepeaterTable(
            schema, "/BCReportInformation", ThreeColumns, new RepeaterTableOptions(), nextId: () => 1));

        Assert.Contains("system data item", ex.Message);
    }

    // ---- negative: column validation ----

    [Fact]
    public void BuildRepeaterTable_column_that_is_a_nested_data_item_throws_ArgumentException()
    {
        var schema = SchemaProvider.FromLayout(Corpus.Path(Corpus.SalesInvoice));

        // AssemblyLine is itself a nested data item under Line in the schema (it has its own child
        // columns - Description_AssemblyLine, LineNo_AssemblyLine, etc. - confirmed against the real
        // dataset custom XML part), not a leaf column, regardless of whether this layout's own document.xml
        // happens to bind a control to it.
        var ex = Assert.Throws<ArgumentException>(() => SdtFactory.BuildRepeaterTable(
            schema, LineItemPath, new[] { "ItemNo_Line", "AssemblyLine" }, new RepeaterTableOptions(),
            nextId: () => 1));

        Assert.Contains("nested repeating data item", ex.Message);
    }

    [Fact]
    public void BuildRepeaterTable_column_not_found_on_the_dataItem_throws_ArgumentException()
    {
        var schema = SchemaProvider.FromLayout(Corpus.Path(Corpus.SalesInvoice));

        var ex = Assert.Throws<ArgumentException>(() => SdtFactory.BuildRepeaterTable(
            schema, LineItemPath, new[] { "ThisColumnDoesNotExistAnywhere" }, new RepeaterTableOptions(),
            nextId: () => 1));

        Assert.Contains("not a leaf column", ex.Message);
    }

    [Fact]
    public void BuildRepeaterTable_zero_columns_throws_ArgumentException()
    {
        var schema = SchemaProvider.FromLayout(Corpus.Path(Corpus.SalesInvoice));

        var ex = Assert.Throws<ArgumentException>(() => SdtFactory.BuildRepeaterTable(
            schema, LineItemPath, Array.Empty<string>(), new RepeaterTableOptions(), nextId: () => 1));

        Assert.Contains("At least one column", ex.Message);
    }

    // ---- THE KEY PROOF: merge round-trip ----

    [Fact]
    public void InsertRepeaterTable_new_repeater_expands_via_MergeEngine_with_distinct_filled_rows()
    {
        var editedPath = CopyOfCorpus(Corpus.SalesInvoice);
        var baselinePath = CopyOfCorpus(Corpus.SalesInvoice); // unedited copy, for a relative baseline.
        var baselineMergedPath =
            Path.Combine(Path.GetTempPath(), $"bcwl-repeatertable-merge-baseline-{Guid.NewGuid():N}.docx");
        var editedMergedPath =
            Path.Combine(Path.GetTempPath(), $"bcwl-repeatertable-merge-edited-{Guid.NewGuid():N}.docx");

        try
        {
            var insertResult = InsertThreeColumnTable(editedPath);

            var options = new MergeOptions { Seed = 424242, Rows = 3 };
            var baseline = MergeEngine.Merge(baselinePath, baselineMergedPath, options);
            var edited = MergeEngine.Merge(editedPath, editedMergedPath, options);

            // Exactly one new repeater expansion generating exactly `Rows` new rows relative to the
            // unedited baseline - our new repeater, and nothing else about the corpus, changed.
            Assert.Equal(baseline.Stats.RepeatersExpanded + 1, edited.Stats.RepeatersExpanded);
            Assert.Equal(baseline.Stats.RowsGenerated + options.Rows, edited.Stats.RowsGenerated);
            Assert.Equal(0, edited.Stats.Unresolved);
            Assert.DoesNotContain(edited.Warnings, w => w.Kind == "unresolved-binding" || w.Kind == "xpath-fallback");

            using var mergedDoc = WordprocessingDocument.Open(editedMergedPath, false);
            var repeaterRow = mergedDoc.MainDocumentPart!.Document!.Descendants<SdtRow>()
                .Single(s => ReadId(s) == insertResult.ControlId);

            var itemRows = repeaterRow.Descendants<SdtRow>().Where(IsRepeaterItemSdt).ToList();
            Assert.Equal(options.Rows, itemRows.Count);

            // Each cloned row's 3 cell values are non-placeholder and distinct from every other row - proof
            // of genuine per-row re-anchoring (not the template's own value repeated verbatim).
            var rowValues = itemRows
                .Select(r => r.Descendants<Text>().Select(t => t.Text).Where(t => !string.IsNullOrEmpty(t)).ToList())
                .ToList();

            Assert.All(rowValues, values =>
            {
                Assert.Equal(3, values.Count);
                Assert.All(values, v => Assert.DoesNotContain('«', v)); // '«' unresolved-binding marker
            });
            Assert.Equal(options.Rows, rowValues.Select(v => string.Join("|", v)).Distinct().Count());

            AssertNoOpenXmlErrors(mergedDoc);
        }
        finally
        {
            File.Delete(editedPath);
            File.Delete(baselinePath);
            File.Delete(baselineMergedPath);
            File.Delete(editedMergedPath);
        }
    }

    // ---- Follow-up review (post-Phase-4.3): a repeater TABLE is v1-scoped to the body ONLY - unlike
    // insert_field/insert_label, which fully support layoutPart='header'/'footer', insert_repeater_table
    // must reject Location.Part != Body outright (repeaters in headers/footers are explicitly deferred -
    // GitHub issue #10). ----

    [Fact]
    public void InsertRepeaterTable_targeting_LayoutPart_Header_is_rejected_and_leaves_the_document_unmodified()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var sdtCountBefore = doc.MainDocumentPart!.Document!.Descendants<SdtElement>().Count();

            var ex = Assert.Throws<ArgumentException>(() => LayoutEditor.InsertRepeaterTable(
                doc, LineItemPath, ThreeColumns, new Location { Type = LocationKind.DocumentEnd, Part = LayoutPart.Header },
                new RepeaterTableOptions()));

            Assert.Equal("location", ex.ParamName);
            Assert.Contains("repeaters in headers/footers", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("insert_field", ex.Message);

            // Rejected before any schema/id/table-building work touches the document at all - not just the
            // main body, but every header part too (nothing was cloned into any of them either).
            Assert.Equal(sdtCountBefore, doc.MainDocumentPart!.Document!.Descendants<SdtElement>().Count());
            foreach (var header in doc.MainDocumentPart.HeaderParts)
            {
                Assert.DoesNotContain(header.Header!.Descendants<SdtElement>(), IsRepeaterItemSdt);
            }
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertRepeaterTable_targeting_LayoutPart_Footer_is_also_rejected()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);

            var ex = Assert.Throws<ArgumentException>(() => LayoutEditor.InsertRepeaterTable(
                doc, LineItemPath, ThreeColumns, new Location { Type = LocationKind.DocumentEnd, Part = LayoutPart.Footer },
                new RepeaterTableOptions()));

            Assert.Equal("location", ex.ParamName);
            Assert.Contains("repeaters in headers/footers", ex.Message, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertRepeaterTable_via_the_tool_surface_with_layoutPart_header_returns_invalid_argument_and_leaves_the_file_untouched()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var before = File.ReadAllBytes(path);

            var response = TableTools.InsertRepeaterTable(
                path, LineItemPath, string.Join(",", ThreeColumns), "documentEnd", layoutPart: "header");

            Assert.False(response.Ok);
            Assert.NotNull(response.Error);
            Assert.Equal("invalid_argument", response.Error!.Code);
            Assert.False(string.IsNullOrWhiteSpace(response.Error.Hint));
            Assert.Equal(before, File.ReadAllBytes(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertRepeaterTable_targeting_LayoutPart_Body_explicitly_still_works_exactly_like_the_default()
    {
        // Regression guard: the new v1-scope rejection must not accidentally reject the (only) supported
        // value too - explicit Body behaves identically to omitting Part entirely.
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);

            var result = LayoutEditor.InsertRepeaterTable(
                doc, LineItemPath, ThreeColumns, new Location { Type = LocationKind.DocumentEnd, Part = LayoutPart.Body },
                new RepeaterTableOptions());

            Assert.Equal("document.xml", result.Part);
            Assert.Equal("Repeater", result.Kind);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- regression: two tables inserted back-to-back must never end up directly adjacent ----

    [Fact]
    public void InsertRepeaterTable_twice_at_documentEnd_never_produces_adjacent_tables_and_both_repeaters_work()
    {
        // Word requires a paragraph between two top-level tables - without one, two adjacent <w:tbl><w:tbl>
        // elements get silently MERGED into a single table by Word on load, corrupting both repeaters. This
        // proves InsertionAnchor.InsertBlock's trailing-paragraph guarantee holds across repeated inserts at
        // the SAME location, not just a single one.
        var editedPath = CopyOfCorpus(Corpus.SalesInvoice);
        var baselinePath = CopyOfCorpus(Corpus.SalesInvoice);
        var baselineMergedPath =
            Path.Combine(Path.GetTempPath(), $"bcwl-repeatertable-twice-baseline-{Guid.NewGuid():N}.docx");
        var editedMergedPath =
            Path.Combine(Path.GetTempPath(), $"bcwl-repeatertable-twice-edited-{Guid.NewGuid():N}.docx");

        try
        {
            var first = InsertThreeColumnTable(editedPath);
            var second = InsertThreeColumnTable(editedPath);
            Assert.NotEqual(first.ControlId, second.ControlId);

            using (var reopened = WordprocessingDocument.Open(editedPath, false))
            {
                var body = reopened.MainDocumentPart!.Document!.Body!;

                // No Table anywhere in the body may be directly followed by another Table (document-wide,
                // not just checked around our two insertions) - a paragraph must always separate them.
                var bodyChildren = body.ChildElements.ToList();
                for (var i = 0; i < bodyChildren.Count - 1; i++)
                {
                    if (bodyChildren[i] is Table)
                    {
                        Assert.IsNotType<Table>(bodyChildren[i + 1]);
                    }
                }

                var openXmlErrors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened).ToList();
                Assert.Empty(openXmlErrors);

                // LayoutReader sees TWO distinct new repeaters (not one merged/corrupted control).
                var inventory = LayoutReader.Read(reopened);
                var ourRepeaters = inventory.Controls
                    .Where(c => c.Kind == ControlKind.Repeater
                        && (c.SdtId == first.ControlId || c.SdtId == second.ControlId))
                    .ToList();
                Assert.Equal(2, ourRepeaters.Count);
            }

            var options = new MergeOptions { Seed = 13131, Rows = 2 };
            var baseline = MergeEngine.Merge(baselinePath, baselineMergedPath, options);
            var edited = MergeEngine.Merge(editedPath, editedMergedPath, options);

            // Both new repeaters expand independently - +2 relative to the unedited baseline.
            Assert.Equal(baseline.Stats.RepeatersExpanded + 2, edited.Stats.RepeatersExpanded);
            Assert.Equal(baseline.Stats.RowsGenerated + (2 * options.Rows), edited.Stats.RowsGenerated);
            Assert.Equal(0, edited.Stats.Unresolved);

            using var mergedDoc = WordprocessingDocument.Open(editedMergedPath, false);
            foreach (var controlId in new[] { first.ControlId, second.ControlId })
            {
                var repeaterRow = mergedDoc.MainDocumentPart!.Document!.Descendants<SdtRow>()
                    .Single(s => ReadId(s) == controlId);
                var itemRows = repeaterRow.Descendants<SdtRow>().Where(IsRepeaterItemSdt).ToList();
                Assert.Equal(options.Rows, itemRows.Count);
            }

            AssertNoOpenXmlErrors(mergedDoc);
        }
        finally
        {
            File.Delete(editedPath);
            File.Delete(baselinePath);
            File.Delete(baselineMergedPath);
            File.Delete(editedMergedPath);
        }
    }

    // ---- tableStyle: the w:tblStyle reference against the layout's own styles part (issue #3 knock-on) ----

    [Fact]
    public void TableStyle_TableGrid_resolves_against_a_blank_created_layouts_scaffolded_styles()
    {
        // Before DefaultStylesScaffold, a from-scratch layout had no styles part at all, so
        // tableStyle='TableGrid' - the tool description's own example - wrote a w:tblStyle reference to a
        // style that did not exist and silently did nothing. The blank build now defines TableGrid, so the
        // emitted reference must both be present on the new table and resolve cleanly per quick
        // validation's table-style-resolves check.
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-repeatertable-styled-{Guid.NewGuid():N}.docx");
        try
        {
            LayoutBuilder.Create(Corpus.Path(Corpus.SalesInvoice), path);

            EditResult result;
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                result = LayoutEditor.InsertRepeaterTable(
                    doc, LineItemPath, ThreeColumns, new Location { Type = LocationKind.DocumentEnd },
                    new RepeaterTableOptions { TableStyle = "TableGrid" });
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var repeater = reopened.MainDocumentPart!.Document!.Descendants<SdtElement>()
                .Single(s => ReadId(s) == result.ControlId);
            var table = repeater.Ancestors<Table>().Single();
            Assert.Equal("TableGrid", table.GetFirstChild<TableProperties>()!.GetFirstChild<TableStyle>()!.Val!.Value);

            var quick = LayoutValidator.Quick(reopened);
            Assert.Empty(quick.Findings.Where(f => f.Check == "table-style-resolves"));
            Assert.True(quick.Passed);
            AssertNoOpenXmlErrors(reopened);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void TableStyle_naming_an_undefined_style_is_flagged_as_a_warning_but_the_resolving_sibling_is_not()
    {
        // The defect fixture (a dangling w:tblStyle) paired with a valid sibling (a resolving one) in the
        // SAME document, per the validator-tests non-tautology rule. A dangling reference is a WARNING,
        // never an error: the layout still renders, the reference just silently does nothing.
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-repeatertable-badstyle-{Guid.NewGuid():N}.docx");
        try
        {
            LayoutBuilder.Create(Corpus.Path(Corpus.SalesInvoice), path);

            using (var doc = WordprocessingDocument.Open(path, true))
            {
                LayoutEditor.InsertRepeaterTable(
                    doc, LineItemPath, ThreeColumns, new Location { Type = LocationKind.DocumentEnd },
                    new RepeaterTableOptions { TableStyle = "TableGrid" });
                LayoutEditor.InsertRepeaterTable(
                    doc, LineItemPath, ThreeColumns, new Location { Type = LocationKind.DocumentEnd },
                    new RepeaterTableOptions { TableStyle = "NoSuchStyle" });
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var quick = LayoutValidator.Quick(reopened);

            var styleFindings = quick.Findings.Where(f => f.Check == "table-style-resolves").ToList();
            var finding = Assert.Single(styleFindings);
            Assert.Equal(FindingSeverity.Warning, finding.Severity);
            Assert.Contains("NoSuchStyle", finding.Message);
            Assert.Equal("document.xml", finding.Location);

            // A warning-only finding must never fail quick validation on its own.
            Assert.True(quick.Passed);
            Assert.Equal(0, quick.ErrorCount);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
