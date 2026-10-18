using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;
using BcWordLayout.McpHost;
using BcWordLayout.McpHost.Tools;
using DocumentFormat.OpenXml.Packaging;

namespace BcWordLayout.Tests;

/// <summary>
/// Drives every <c>case</c> label in
/// <c>ToolGuards.InvalidArgumentHint</c>'s switch end to end, asserting a hint fragment DISTINCTIVE to that
/// branch (never a fragment every hint shares) rather than merely that a hint is present. This is the test
/// signal the hint's own doc comment now promises: renaming/retiring a <c>nameof(...)</c> at any of the
/// throw-site files that switch keys off (<c>ToolGuards</c>, <c>SdtFactory</c>, <c>LayoutEditor</c>,
/// <c>Location.Validate</c>, <c>LocationResolver</c>, <c>TableStructureEditor</c>, <c>CellTextEditor</c>,
/// <c>LayoutBuilder</c>, <c>LayoutRefresher</c>) now breaks a SPECIFIC named test below instead of silently
/// rerouting that failure to the generic fallback with no compile/test signal.
/// <para>
/// Building this class's own cross-map (every switch key vs. every real throw site) surfaced two classes of
/// pre-existing drift, both fixed alongside this file (see <c>InvalidArgumentHint</c>'s own remarks for the
/// full account): DEAD keys removed (<c>"row"</c>/<c>"col"</c>, <c>"field"</c>/<c>"label"</c>/<c>"dataItem"</c>
/// — no throw site anywhere ever produces those exact <see cref="System.ArgumentException.ParamName"/>
/// values) and previously-UNMAPPED real ParamNames given their own case (<c>"schemaSource"</c>/
/// <c>"outputPath"</c>/<c>"newSchemaSource"</c> — live, tool-reachable throw sites that used to silently fall
/// back to the generic hint).
/// </para>
/// Complements, rather than duplicates, <see cref="ErrorHintTests"/> (which covers a similar but
/// non-exhaustive set of failure shapes, plus <c>file_not_found</c>/<c>invalid_layout</c>/<c>not_found</c>
/// coverage this file does not repeat) — this file's job is exhaustive KEY coverage, one representative
/// throw site per key, not per-tool robustness.
/// </summary>
public class InvalidArgumentHintCoverageTests
{
    private static void AssertActionableFailure(ToolResponse response, string expectedCode, string hintMustMention)
    {
        Assert.False(response.Ok);
        Assert.Null(response.Data);
        Assert.NotNull(response.Error);

        var error = response.Error!;
        Assert.Equal(expectedCode, error.Code);
        Assert.False(string.IsNullOrWhiteSpace(error.Message));
        Assert.False(string.IsNullOrWhiteSpace(error.Hint));
        Assert.Contains(hintMustMention, error.Hint, StringComparison.OrdinalIgnoreCase);
    }

    private static string CopyOfCorpus(string corpusFile)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-hintcov-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    private static string NewOutputPath() =>
        Path.Combine(Path.GetTempPath(), $"bcwl-hintcov-out-{Guid.NewGuid():N}.docx");

    // ---- "locationtype" — ToolGuards.BuildLocation / TryParseLocationKind ----

    [Fact]
    public void InsertField_bad_locationType_hint_names_the_valid_locationTypes()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = EditTools.InsertField(path, "/Header/CustomerAddress1", "not-a-real-location");
            AssertActionableFailure(response, "invalid_argument", "locationType must be one of");
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- "layoutpart" — ToolGuards.BuildLocation / TryParseLayoutPart ----

    [Fact]
    public void InsertField_bad_layoutPart_hint_names_the_valid_layoutParts()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = EditTools.InsertField(
                path, "/Header/CustomerAddress1", "documentEnd", layoutPart: "bogus");
            AssertActionableFailure(
                response, "invalid_argument", "partName (optional) only applies when layoutPart is");
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- "location" — LayoutEditor.InsertRepeaterTable's v1 body-only scope check ----

    [Fact]
    public void InsertRepeaterTable_targeting_a_header_hint_names_v1_scope()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = TableTools.InsertRepeaterTable(
                path, "/Header/Line", "ItemNo_Line", "documentEnd", layoutPart: "header");
            AssertActionableFailure(response, "invalid_argument", "repeating section");
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- "controlid" — Location.Validate (AfterControl requires ControlId) ----

    [Fact]
    public void InsertField_afterControl_without_controlId_hint_names_controlId_requirement()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = EditTools.InsertField(path, "/Header/CustomerAddress1", "afterControl");
            AssertActionableFailure(response, "invalid_argument", "afterControl requires controlId");
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- "tableindex" — Location.Validate (TableCell requires TableIndex/Row/Col); ParamName is ALWAYS
    // "TableIndex" here even when Row or Col specifically is the missing/negative field (see the dead "row"/
    // "col" keys removed from the switch — no throw site anywhere ever reports those as ParamName). ----

    [Fact]
    public void InsertField_tableCell_missing_row_and_col_hint_names_tableCell_requirements()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            // tableIndex supplied, row/col omitted — still incomplete for tableCell addressing.
            var response = EditTools.InsertField(path, "/Header/CustomerAddress1", "tableCell", tableIndex: 0);
            AssertActionableFailure(
                response, "invalid_argument", "requires non-negative tableIndex, row, and col");
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- "searchtext" — Location.Validate (AtText requires SearchText) ----

    [Fact]
    public void InsertField_atText_without_searchText_hint_names_searchText_requirement()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = EditTools.InsertField(path, "/Header/CustomerAddress1", "atText");
            AssertActionableFailure(response, "invalid_argument", "atText requires a non-empty searchText substring");
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- "datasetpath" / "dataitempath" — SdtFactory.ValidatedSegments/ValidateLeafColumn/ResolveRepeaterDataItem.
    // (the dead "field"/"label"/"dataitem" sub-keys were removed: insert_field/insert_label/
    // insert_repeater_table's OWN parameter names are never the reported ParamName — validation always
    // happens one layer down in SdtFactory under "datasetPath"/"dataItemPath".) ----

    [Fact]
    public void InsertField_dataset_path_that_does_not_resolve_hint_names_the_field_vs_label_rule()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = EditTools.InsertField(path, "/Header/ThisFieldDoesNotExistAnywhere", "documentEnd");
            AssertActionableFailure(response, "invalid_argument", "insert_field requires the reverse");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertRepeaterTable_with_a_leaf_column_as_dataItem_hint_names_the_data_item_rule()
    {
        // Pins the arm's OTHER ParamName, "dataItemPath" (SdtFactory.ResolveRepeaterDataItem) — without
        // this, renaming that parameter would silently degrade insert_repeater_table's bad-dataItem
        // failures to the generic fallback while the datasetPath test above stayed green.
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = TableTools.InsertRepeaterTable(
                path, "/Header/Line/ItemNo_Line", "ItemNo_Line", "documentEnd");
            AssertActionableFailure(response, "invalid_argument", "must be a repeating, non-system DATA ITEM");
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- "columns" / "column" — TableTools' own empty-columns pre-check (ParamName "columns") and
    // SdtFactory.ValidateColumnOfItem's per-entry check (ParamName "column") — same combined arm/hint. ----

    [Theory]
    [InlineData("   ")]
    [InlineData("ThisColumnDoesNotExistAnywhere")]
    public void InsertRepeaterTable_bad_columns_hint_names_leaf_columns_of_the_dataItem(string columns)
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = TableTools.InsertRepeaterTable(path, "/Header/Line", columns, "documentEnd");
            AssertActionableFailure(response, "invalid_argument", "belonging to the given dataItem");
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- "columnwidths" / "options" — TableTools.ParseColumnWidths' bad-integer-format check (ParamName
    // "columnWidths") and SdtFactory.BuildRepeaterTable's count-mismatch check (ParamName "options") — same
    // combined arm/hint. ----

    [Fact]
    public void InsertRepeaterTable_columnWidths_bad_integer_format_hint_names_columnWidths_shape()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = TableTools.InsertRepeaterTable(
                path, "/Header/Line", "ItemNo_Line,Description_Line,TransHeaderAmount", "documentEnd",
                columnWidths: "abc,2000,3000");
            AssertActionableFailure(response, "invalid_argument", "omit it entirely for an even default width");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertRepeaterTable_columnWidths_count_mismatch_hint_names_columnWidths_shape()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = TableTools.InsertRepeaterTable(
                path, "/Header/Line", "ItemNo_Line,Description_Line,TransHeaderAmount", "documentEnd",
                columnWidths: "1000,2000");
            AssertActionableFailure(response, "invalid_argument", "omit it entirely for an even default width");
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- "widths" — TableTools.ParseIntListOrThrow, called only from set_column_widths ----

    [Fact]
    public void SetColumnWidths_bad_integer_format_hint_names_set_column_widths_shape()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(rows: 2, cols: 3));
        try
        {
            var response = TableTools.SetColumnWidths(path, tableIndex: 0, widths: "1000,abc,2000");
            AssertActionableFailure(response, "invalid_argument", "set_column_widths' widths must be a non-empty");
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- "cells" — TableTools.ParseRepeaterRowCells' span range, and the domain's span-sum check, shared by
    // insert_table_row/insert_repeater_row/insert_subtotal_row ----

    [Theory]
    [InlineData("-1:-,4:-")]
    [InlineData("2:-")]
    public void InsertTableRow_bad_cells_hint_names_the_cells_dsl(string cells)
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(rows: 2, cols: 3));
        try
        {
            var response = TableTools.InsertTableRow(path, tableIndex: 0, cells: cells);
            AssertActionableFailure(response, "invalid_argument", "cells is a comma-separated list of cell specs");
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- "atcolumn" — TableStructureEditor.InsertColumn's range / spanned-content-cell checks ----

    [Fact]
    public void InsertColumn_atColumn_out_of_range_hint_names_the_valid_grid_positions()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(rows: 2, cols: 3));
        try
        {
            var response = TableTools.InsertColumn(path, tableIndex: 0, mode: "plainText", atColumn: 99);
            AssertActionableFailure(response, "invalid_argument", "0-based GRID position");
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- "width" — TableStructureEditor.InsertColumn's width must leave room for the existing columns ----

    [Fact]
    public void InsertColumn_width_as_wide_as_the_table_hint_names_the_kept_total()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(rows: 2, cols: 3));
        try
        {
            var response = TableTools.InsertColumn(path, tableIndex: 0, mode: "plainText", width: 6000);
            AssertActionableFailure(response, "invalid_argument", "smaller than the table's current total width");
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- "gridcolumn" — TableStructureEditor.RemoveColumn's range check ----

    [Fact]
    public void RemoveColumn_gridColumn_out_of_range_hint_names_grid_column_addressing()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(rows: 2, cols: 3));
        try
        {
            var response = TableTools.RemoveColumn(path, tableIndex: 0, column: 99);
            AssertActionableFailure(response, "invalid_argument", "0-based GRID column index");
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- "fromcolumn" / "tocolumn" — TableStructureEditor.MergeCells' range check (ParamName "fromColumn")
    // and its absorbed-bound-cell check (ParamName "toColumn") — same combined arm/hint. ----

    [Fact]
    public void MergeCells_range_out_of_bounds_hint_names_the_absorbed_cell_rule()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(rows: 2, cols: 3));
        try
        {
            var response = TableTools.MergeCells(path, tableIndex: 0, row: 0, fromColumn: 5, toColumn: 6);
            AssertActionableFailure(
                response, "invalid_argument", "absorbed cell that holds a bound control is rejected");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void MergeCells_absorbing_a_bound_cell_hint_names_the_absorbed_cell_rule()
    {
        var body =
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"1000\"/><w:gridCol w:w=\"1000\"/></w:tblGrid>"
            + "<w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc>"
            + "<w:tc><w:tcPr/><w:sdt><w:sdtPr><w:id w:val=\"5\"/></w:sdtPr>"
            + "<w:sdtContent><w:p><w:r><w:t>b</w:t></w:r></w:p></w:sdtContent></w:sdt></w:tc></w:tr></w:tbl>";
        var path = SyntheticLayout.Create(body);
        try
        {
            var response = TableTools.MergeCells(path, tableIndex: 0, row: 0, fromColumn: 0, toColumn: 1);
            AssertActionableFailure(
                response, "invalid_argument", "absorbed cell that holds a bound control is rejected");
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- "cellindex" — TableStructureEditor.SplitCell's range/unspanned check ----

    [Fact]
    public void SplitCells_cellIndex_out_of_range_hint_names_the_spanned_cell_rule()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(rows: 1, cols: 3));
        try
        {
            var response = TableTools.SplitCells(path, tableIndex: 0, row: 0, cellIndex: 99);
            AssertActionableFailure(response, "invalid_argument", "gridSpan > 1");
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- "mode" — TableTools.ParseInsertColumnModeOrThrow ----

    [Fact]
    public void InsertColumn_bad_mode_hint_names_the_valid_modes()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(rows: 2, cols: 3));
        try
        {
            var response = TableTools.InsertColumn(path, tableIndex: 0, mode: "bogus");
            AssertActionableFailure(response, "invalid_argument", "'field', 'label', or 'plainText'");
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- "format" — RunFormatter.Prepare (a member out of range, malformed, or an undefined style) ----

    [Fact]
    public void InsertText_bad_format_member_hint_lists_every_member_and_its_values()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = EditTools.InsertText(
                path, "x", "documentEnd", format: new TextFormatArgument { Color = "blue" });
            AssertActionableFailure(response, "invalid_argument", "six hex digits RRGGBB");
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- "target" — EditTools.ParseTextFormatTargetOrThrow and TextFormatEditor's target-shape check ----

    [Fact]
    public void SetTextFormat_target_missing_its_arguments_hint_lists_each_target_and_what_it_needs()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = EditTools.SetTextFormat(
                path, "cell", tableIndex: 0, format: new TextFormatArgument { Bold = true });
            AssertActionableFailure(response, "invalid_argument", "'row' (with tableIndex and row");
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- "alignment" — CellTextEditor.ParseAlignment, from set_cell_text ----

    [Fact]
    public void SetCellText_bad_alignment_hint_names_the_valid_alignments()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(rows: 1, cols: 2));
        try
        {
            var response = EditTools.SetCellText(path, 0, 0, 0, "x", alignment: "middle");
            AssertActionableFailure(response, "invalid_argument", "'left', 'center' or 'right'");
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- "look" — TableTools.ParseTableBorderLookOrThrow ----

    [Fact]
    public void InsertRepeaterTable_bad_look_hint_names_the_valid_looks()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = TableTools.InsertRepeaterTable(
                path, "/Header/Line", "ItemNo_Line", "documentEnd", look: "fancy");
            AssertActionableFailure(response, "invalid_argument", "'bc'");
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- "edges" / "style" / "size" — TableTools.ParseCellBorderOptionsOrThrow + TableStructureEditor.SetCellBorders ----

    [Fact]
    public void SetCellBorders_bad_edges_hint_names_the_valid_edges()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(rows: 2, cols: 3));
        try
        {
            var response = TableTools.SetCellBorders(path, tableIndex: 0, row: 0, edges: "sideways");
            AssertActionableFailure(response, "invalid_argument", "'bottom'");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void SetCellBorders_bad_size_hint_names_the_eighths_of_a_point_range()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(rows: 2, cols: 3));
        try
        {
            var response = TableTools.SetCellBorders(path, tableIndex: 0, row: 0, edges: "top", size: 0);
            AssertActionableFailure(response, "invalid_argument", "eighths of a point");
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- "datapath" — TableStructureEditor.InsertColumn's mode='field'/'label' requires dataPath ----

    [Fact]
    public void InsertColumn_field_mode_missing_dataPath_hint_names_insert_column_dataPath()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(rows: 2, cols: 3));
        try
        {
            var response = TableTools.InsertColumn(path, tableIndex: 0, mode: "field");
            AssertActionableFailure(response, "invalid_argument", "insert_column's dataPath must be a full dataset path");
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- "headertext" / "headerlabelpath" — TableStructureEditor.PlanInsertColumnRows (no header row to
    // hold them) and ValidateInsertColumnPaths (headerLabelPath does not resolve) ----

    [Fact]
    public void InsertColumn_headerText_on_a_table_with_no_header_row_hint_names_the_header_rows()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(rows: 2, cols: 3));
        try
        {
            var response = TableTools.InsertColumn(path, tableIndex: 0, mode: "plainText", headerText: "Notes");
            AssertActionableFailure(response, "invalid_argument", "fill the new column's cell in the table's HEADER");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertColumn_unresolvable_headerLabelPath_hint_names_the_header_rows()
    {
        var path = CopyOfCorpus(Corpus.StandardSalesQuote);
        try
        {
            var response = TableTools.InsertColumn(
                path, tableIndex: 1, mode: "plainText", headerLabelPath: "/Header/NoSuchCaption_Lbl");
            AssertActionableFailure(response, "invalid_argument", "fill the new column's cell in the table's HEADER");
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- "firstcontrolid" / "lastcontrolid" / "wholebody" — LayoutEditor.WrapInRepeater's range ----

    /// <summary>A blank quote layout with two work-description fields, each in its own body paragraph.</summary>
    private static (string Path, int First, int Second) LayoutWithTwoWorkDescriptionFields()
    {
        var path = NewOutputPath();
        LayoutBuilder.Create(Corpus.Path(Corpus.StandardSalesQuote), path);
        using var doc = WordprocessingDocument.Open(path, true);
        int Insert() => LayoutEditor.InsertField(doc, "/Header/WorkDescriptionLines/WorkDescriptionLine",
            new Location { Type = LocationKind.DocumentEnd }).ControlId;
        var first = Insert();
        var second = Insert();
        doc.MainDocumentPart!.Document!.Save();
        return (path, first, second);
    }

    [Fact]
    public void WrapInRepeater_with_no_range_hint_names_the_range_rule()
    {
        var (path, _, _) = LayoutWithTwoWorkDescriptionFields();
        try
        {
            var response = EditTools.WrapInRepeater(path, "/Header/WorkDescriptionLines");
            AssertActionableFailure(response, "invalid_argument", "firstControlId through lastControlId");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void WrapInRepeater_with_controls_in_the_wrong_order_hint_names_the_range_rule()
    {
        var (path, first, second) = LayoutWithTwoWorkDescriptionFields();
        try
        {
            var response = EditTools.WrapInRepeater(path, "/Header/WorkDescriptionLines", firstControlId: second, lastControlId: first);
            AssertActionableFailure(response, "invalid_argument", "earlier in the document than the last");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void WrapInRepeater_with_ids_and_wholeBody_hint_names_the_range_rule()
    {
        var (path, first, _) = LayoutWithTwoWorkDescriptionFields();
        try
        {
            var response = EditTools.WrapInRepeater(path, "/Header/WorkDescriptionLines", firstControlId: first, wholeBody: true);
            AssertActionableFailure(response, "invalid_argument", "wholeBody=true and no ids");
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- "keeptext" — LayoutEditor.RemoveControl (keepText=true on a repeater) ----

    [Fact]
    public void RemoveControl_keepText_on_a_repeater_hint_names_keepText_false()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var info = Assert.IsType<LayoutInfoDto>(ReadTools.GetLayoutInfo(path).Data);
            // A ROW-level repeater: keepText=true unwraps a block-level one (wrap_in_repeater's undo, #31).
            var repeater = info.Controls.First(c => c.Kind == "Repeater" && c.Level == "row" && c.SdtId.HasValue);

            var response = EditTools.RemoveControl(path, repeater.SdtId!.Value, keepText: true);

            AssertActionableFailure(response, "invalid_argument", "keepText=false");
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- "schema" — SdtFactory's storeItemID guard (schema.Report.StoreItemId null/empty). No tool call
    // reaches it any more. It used to be reached by every bound insert on a layout whose dataset part has no
    // itemProps - 4 of the 10 corpus layouts, stock ones included - with a hint naming internal C# methods
    // (GitHub issue #93). Every tool that builds a bound control now takes its schema from
    // SchemaProvider.FromLayoutForBinding, which supplies the part's itemID or the storeItemID the layout's
    // existing bindings agree on, and otherwise refuses with invalid_layout (StoreItemIdFallbackTests covers
    // both). So, like the other defensive keys, it is pinned through the public API its throw site lives on:
    // SdtFactory handed a schema parsed from a standalone schema .xml (which has no storeItemID), run inside
    // the same ToolGuards.Guard every tool body runs in. ----

    [Fact]
    public void SdtFactory_given_a_schema_with_no_storeItemId_hint_reports_a_tool_bug()
    {
        var xml = Path.Combine(Path.GetTempPath(), $"bcwl-hintcov-schema-{Guid.NewGuid():N}.xml");
        File.WriteAllText(xml, SyntheticLayout.DatasetWithLineXml);
        try
        {
            var response = ToolGuards.Guard(
                () => ToolResponse.Success(
                    BcWordLayout.Domain.SdtFactory.BuildField(
                        BcWordLayout.Domain.SchemaProvider.FromSchemaXml(xml), "/Header/CompanyName")),
                xml);

            AssertActionableFailure(response, "invalid_argument", "resolves one from the layout itself");
            Assert.DoesNotContain("SchemaProvider", response.Error!.Hint, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(xml);
        }
    }

    // ---- "schemasource" — LayoutBuilder.Create (schemaSource empty) ----

    [Fact]
    public void CreateLayout_empty_schemaSource_hint_names_create_layout_schemaSource()
    {
        var response = LifecycleTools.CreateLayout("   ", NewOutputPath());
        AssertActionableFailure(response, "invalid_argument", "create_layout's schemaSource must be a non-empty");
    }

    // ---- "outputpath" — LayoutBuilder.Create (outputPath empty) ----

    [Fact]
    public void CreateLayout_empty_outputPath_hint_names_the_output_path_requirement()
    {
        var response = LifecycleTools.CreateLayout(Corpus.Path(Corpus.SalesInvoice), "   ");
        AssertActionableFailure(response, "invalid_argument", "outputPath is where the new");
    }

    // ---- "templatepath" — LayoutBuilder.Create (a macro-enabled templatePath, GitHub issue #77) ----

    [Fact]
    public void CreateLayout_macro_enabled_templatePath_hint_names_create_layout_templatePath()
    {
        var templatePath = SyntheticLayout.CreateWordPackage(
            DocumentFormat.OpenXml.WordprocessingDocumentType.MacroEnabledTemplate, ".dotm");
        var outputPath = NewOutputPath();
        try
        {
            var response = LifecycleTools.CreateLayout(Corpus.Path(Corpus.SalesInvoice), outputPath, templatePath);
            AssertActionableFailure(response, "invalid_argument", "create_layout's templatePath");
            Assert.False(File.Exists(outputPath));
        }
        finally
        {
            File.Delete(templatePath);
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
        }
    }

    // ---- "newschemasource" — LayoutRefresher.Refresh (newSchemaSource empty) ----

    [Fact]
    public void RefreshXmlPart_empty_newSchemaSource_hint_names_refresh_xml_part_newSchemaSource()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = LifecycleTools.RefreshXmlPart(path, "   ");
            AssertActionableFailure(
                response, "invalid_argument", "refresh_xml_part's newSchemaSource must be a non-empty");
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- "dataoverridespath" — SampleDataGenerator.LoadOverrides (an unusable data file, GitHub issue #95) ----

    [Fact]
    public void Merge_with_a_wrong_shaped_dataOverridesPath_hint_names_the_two_accepted_shapes()
    {
        // Driven through MergeEngine.Merge - the call preview_layout makes - inside the same Guard every tool
        // uses, rather than through preview_layout itself, so this class need not join the
        // preview-converter-seam collection; CallerFileErrorTests covers the tool end to end.
        var overrides = Path.Combine(Path.GetTempPath(), $"bcwl-hintcov-overrides-{Guid.NewGuid():N}.xml");
        var outputPath = NewOutputPath();
        File.WriteAllText(overrides, "<?xml version=\"1.0\"?><SomethingElse/>");
        try
        {
            var response = ToolGuards.Guard(() =>
            {
                BcWordLayout.Merge.MergeEngine.Merge(
                    Corpus.Path(Corpus.SalesInvoice),
                    outputPath,
                    new BcWordLayout.Merge.MergeOptions { DataOverridesPath = overrides });
                return ToolResponse.Success(true);
            }, Corpus.Path(Corpus.SalesInvoice));

            AssertActionableFailure(response, "invalid_argument", "preview_layout's dataOverridesPath must be");
            Assert.False(File.Exists(outputPath));
        }
        finally
        {
            File.Delete(overrides);
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
        }
    }

    // ---- default fallback ("_") — TableStructureEditor.RejectUnsupportedShape throws with NO ParamName at
    // all (single-arg ArgumentException constructor); InvalidArgumentHint(null) must land on the generic
    // fallback rather than any specific case. ----

    [Fact]
    public void SetColumnWidths_on_a_vMerge_table_has_no_paramName_and_falls_back_to_the_generic_hint()
    {
        var body =
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2000\"/><w:gridCol w:w=\"2000\"/></w:tblGrid>"
            + "<w:tr><w:tc><w:tcPr><w:vMerge w:val=\"restart\"/></w:tcPr><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc>"
            + "<w:tc><w:tcPr/><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc></w:tr>"
            + "<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>"
            + "<w:tc><w:tcPr/><w:p><w:r><w:t>c</w:t></w:r></w:p></w:tc></w:tr></w:tbl>";
        var path = SyntheticLayout.Create(body);
        try
        {
            var response = TableTools.SetColumnWidths(path, tableIndex: 0, widths: "2000,2000");
            AssertActionableFailure(response, "invalid_argument", "Check the argument named in the message above");
        }
        finally
        {
            File.Delete(path);
        }
    }
}
