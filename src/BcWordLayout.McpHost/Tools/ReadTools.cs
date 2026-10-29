using System.ComponentModel;
using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;
using BcWordLayout.Merge;
using DocumentFormat.OpenXml.Packaging;
using ModelContextProtocol.Server;
using static BcWordLayout.McpHost.Tools.ToolGuards;

namespace BcWordLayout.McpHost.Tools;

/// <summary>
/// Read-only MCP tools for inspecting a BC Word report layout: <c>get_layout_info</c> (full control/table
/// inventory + quick validation), <c>list_dataset_fields</c> (the report dataset hierarchy), and
/// <c>validate_layout</c> (quick or full validation). None of these MUTATE the layout, but all three
/// still take the SAME per-path lock pair <see cref="EditTools"/>/<see cref="TableTools"/>/
/// <see cref="LifecycleTools"/>'s mutating tools do — via <see cref="ToolGuards.GuardRead"/> rather than
/// <see cref="ToolGuards.GuardMutate{TResult}"/> (a SHORT bounded wait instead of an unbounded one; see
/// <see cref="ToolGuards.GuardRead"/>'s own remarks for why a read must coordinate with the mutating tools'
/// atomic-rename commit at all) — every failure still flows through the shared <see cref="ToolGuards.Guard"/>
/// exception-to-envelope translator, and response DTOs are built via <see cref="ToolGuards"/>'s shared
/// mapping helpers (<see cref="ToolGuards.ToControlDto"/>, <see cref="ToolGuards.ToTableDto"/>,
/// <see cref="ToolGuards.ToDataItemDto"/>, etc.).
/// </summary>
[McpServerToolType]
public static class ReadTools
{
    [McpServerTool(Name = "get_layout_info")]
    [Description("Inspect a BC Word report layout (.docx): returns report name/id, dataset namespace, "
                 + "storeItemId, the full control inventory (kind, alias, tag, xPath, storeItemId, part, "
                 + "sdtId - the control's w:id, the controlId the edit tools take - usesW15Binding, "
                 + "parentRepeaterXPath, and for each control its structural LEVEL - run/block/cell/row/runRuby - "
                 + "plus its tableIndex/rowIndex/colIndex when it sits in a table), a TABLES section "
                 + "describing every table's grid (rowCount, columnCount, gridColumnWidths) and its rows - ONE "
                 + "PER w:tr, so each row of a multi-row repeater item (line, nested detail, subtotal, spacer "
                 + "rows) has its own rowIndex, tagged with the innermost repeater it repeats with "
                 + "(repeaterControlId) - and, per row and "
                 + "cell, which control owns it and the cell's visible text, plus BOTH cell coordinates: "
                 + "colIndex (the physical position in the row - what set_cell_text/merge_cells/split_cells/"
                 + "set_cell_format and tableCell locations take) and gridColumn/gridSpan (where the cell "
                 + "sits on the table grid - what remove_column/insert_column/set_column_widths take; the "
                 + "two differ on any row with spanned cells), with each row's gridBefore/gridAfter skipped "
                 + "columns, a PART DETAILS section giving "
                 + "each OOXML part's kind and header/footer ROLE (default/first/even - plus isDefaultTarget, "
                 + "the part a header/footer edit with no partName lands in) alongside hasTitlePage (whether "
                 + "page 1 renders the 'first'-role parts INSTEAD of the default ones), a quick "
                 + "validation status summary, and TARGET: the Business Central version the layout must work on - "
                 + "bcVersion (absent when unknown), source (app.json - the nearest AL app.json above the layout - "
                 + "environment - BCWL_BC_VERSION - or unknown) and detail (what it was read from, or why there is "
                 + "none). Read target before using a feature that needs a newer BC. IMPORTANT for editing: a control whose level is 'cell' wraps "
                 + "a whole table cell (one column) - e.g. BC header address fields; removing it is safe "
                 + "(remove_control never deletes the cell/column), but use this table detail to understand "
                 + "what a control occupies before editing. Check partDetails BEFORE any header/footer edit: "
                 + "the package-order first part is frequently the even-page or first-page one, not the "
                 + "everyday default.")]
    public static ToolResponse GetLayoutInfo(
        [Description("Absolute path to the .docx layout file.")] string layoutPath)
    {
        return GuardRead(layoutPath, target =>
        {
            if (!File.Exists(layoutPath))
            {
                throw new FileNotFoundException("layoutPath does not point to an existing file.", layoutPath);
            }

            // Open the package once and share the handle across all three passes (schema, inventory,
            // validation) instead of re-parsing the same .docx repeatedly.
            using var doc = LayoutPackage.Open(layoutPath, isEditable: false);
            var tree = SchemaProvider.FromLayout(doc);
            var inventory = LayoutReader.Read(doc);
            var validation = LayoutValidator.Quick(doc, target);

            var controls = inventory.Controls.Select(ToControlDto).ToList();
            var summary = new ControlSummaryDto(
                Field: inventory.Controls.Count(c => c.Kind == ControlKind.Field),
                Label: inventory.Controls.Count(c => c.Kind == ControlKind.Label),
                Repeater: inventory.RepeaterCount,
                Picture: inventory.Controls.Count(c => c.Kind == ControlKind.Picture),
                Unbound: inventory.Controls.Count(c => c.Kind == ControlKind.Unbound),
                Total: inventory.Controls.Count);

            var dto = new LayoutInfoDto(
                Report: ToReportDto(tree.Report),
                Validation: new ValidationSummaryDto("quick", validation.Passed, validation.ErrorCount, validation.WarningCount),
                ControlSummary: summary,
                Parts: inventory.Parts,
                Controls: controls,
                Tables: inventory.Tables.Select(ToTableDto).ToList(),
                PartDetails: inventory.PartDetails.Select(ToPartInfoDto).ToList(),
                HasTitlePage: inventory.HasTitlePage,
                Target: ToTargetDto(target));

            return ToolResponse.Success(dto);
        });
    }

    [McpServerTool(Name = "list_dataset_fields")]
    [Description("List the report dataset hierarchy (data items + columns) from either a .docx layout "
                 + "or a standalone exported schema .xml. Each column is flagged isLabel; when the source "
                 + "is a layout, each column is also flagged bound/unbound based on the layout's controls.")]
    public static ToolResponse ListDatasetFields(
        [Description("Absolute path to a .docx layout OR a standalone schema .xml file.")] string source)
    {
        // The extension dispatch below runs OUTSIDE both guards, so a null source must be rejected here -
        // otherwise source.EndsWith throws a raw NullReferenceException across the tool boundary instead of
        // the structured envelope every other input mistake gets (the sibling read tools reach the same
        // outcome via GuardRead's own path normalization).
        if (source is null)
        {
            return ToolResponse.Failure(
                ErrorCodes.InvalidArgument,
                "source must not be null.",
                "Pass an absolute path to either a .docx layout or a standalone exported schema .xml file.");
        }

        // Only the .docx (layout) branch touches a file the mutating tools could ever be committing to via
        // GuardMutate's atomic rename - a standalone schema .xml is never written by any tool in this
        // codebase, so it has no commit-vs-read race to coordinate against and needs neither half of the
        // edit lock (see ToolGuards.GuardRead's own remarks for what that race actually is).
        if (!source.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
        {
            return Guard(() =>
            {
                var schemaTree = SchemaProvider.FromSchemaXml(source);
                var schemaDto = new DatasetFieldsDto("schema", ToReportDto(schemaTree.Report), ToDataItemDto(schemaTree.Root, null));
                return ToolResponse.Success(schemaDto);
            }, source);
        }

        return GuardRead(source, () =>
        {
            if (!File.Exists(source))
            {
                throw new FileNotFoundException("source does not point to an existing file.", source);
            }

            // Share one open package across the schema parse and the inventory read.
            using var doc = LayoutPackage.Open(source, isEditable: false);
            var tree = SchemaProvider.FromLayout(doc);
            var boundPaths = BuildBoundPaths(LayoutReader.Read(doc));

            var root = ToDataItemDto(tree.Root, boundPaths);
            var dto = new DatasetFieldsDto("layout", ToReportDto(tree.Report), root);
            return ToolResponse.Success(dto);
        });
    }

    [McpServerTool(Name = "validate_layout")]
    [Description("Validate a BC Word layout. level='quick' runs structural + binding checks (OpenXML "
                 + "validity, a document-content-type ERROR when the main part is typed as a Word template or "
                 + "macro-enabled file - which Word refuses to open as a .docx, single BC XML part (with several, "
                 + "the finding lists them and the other checks measure against the one the bindings name by "
                 + "storeItemID; refresh_xml_part and the tools that write a binding refuse such a layout), "
                 + "storeItemID match, binding dataset-namespace match, a "
                 + "binding-prefix-mapping ERROR when a dataset binding's w:prefixMappings does not declare "
                 + "every prefix its XPath uses as the layout's own namespace (a wrong or misspelled namespace, "
                 + "or a prefix the mappings never declare - Word cannot map such a control; a bare URI that "
                 + "names the right namespace is tolerated, as stock layouts ship it), XPath "
                 + "resolution, repeater shape, an attached-template ERROR when the layout links an external "
                 + "attached template (Business Central refuses the upload - UnsupportedExternalContentFields; "
                 + "every tool that writes the layout removes the link), a compatibility-mode warning "
                 + "when a layout that CONTAINS repeaters declares a Word compatibility mode below 15, and "
                 + "a repeater-downgraded ERROR when a control's alias names a whole data item but the "
                 + "control is not a repeating section - the wreckage a Word save leaves when it converts "
                 + "a repeater to a plain rich-text control, which makes the table render one row instead of "
                 + "one per record; and a top-level-repeater-group ERROR when a block-level repeating "
                 + "section repeats the top-level data item without a w:group around its item - Business "
                 + "Central prints such a layout as an empty document while preview and the full merge fill "
                 + "it normally), and a duplicate-control-id WARNING when two content controls share a w:id - the "
                 + "file still opens, but remove_control, afterControl and parentRepeaterId refuse that id "
                 + "rather than act on the first match, and a font-installed WARNING when a font the layout names "
                 + "(on runs, styles, numbering, or the theme's heading/body font where text refers to it) is not "
                 + "one Business Central online has installed and the layout does not embed it - BC prints that "
                 + "text in a substitute font, and a bc-version finding when the layout uses a feature the target "
                 + "Business Central version lacks: a header/footer layout bound to Tenant_Report_Defaults (report "
                 + "themes, BC 29.0 and later) is an ERROR for a lower target and a WARNING for an unknown one. The "
                 + "result's target says which version was used and where it came from. One structural complaint is reported as a WARNING rather than an error: "
                 + "w:rPr/w:w=0, the character scaling Word writes into a content control's placeholder run, which "
                 + "renders correctly in Word and BC and cannot be edited by any tool here. level='full' "
                 + "additionally runs a real dry-run merge (sample data fill + repeater expansion) against "
                 + "a throwaway temp copy and surfaces every merge warning (e.g. unresolved bindings, XPath "
                 + "errors) as a validation finding.")]
    public static ToolResponse ValidateLayout(
        [Description("Absolute path to the .docx layout file.")] string layoutPath,
        [Description("Validation level: 'quick' or 'full'. Default 'quick'.")] [LevelValues] string level = "quick")
    {
        // An explicit JSON null is bound straight through by the SDK; like every other optional enum-like
        // string on this surface, it means the documented default rather than an unknown level.
        level ??= "quick";
        if (!TryParseLevel(level, out var full))
        {
            return Guard(() => ToolResponse.Failure(
                ErrorCodes.InvalidArgument,
                $"Unknown validation level '{level}'.",
                "level must be 'quick' or 'full' (case-insensitive); 'quick' runs fast structural+binding "
                + "checks, 'full' additionally dry-run merges sample data and surfaces merge warnings."), layoutPath);
        }

        return GuardRead(layoutPath, target =>
        {
            // FullValidator.Full also reads layoutPath (via a File.Copy dry-run merge) on top of the quick
            // checks it starts with, so the SAME lock pair covers both levels rather than only the quick one.
            var result = full ? FullValidator.Full(layoutPath, target) : LayoutValidator.Quick(layoutPath, target);

            var dto = new ValidationResultDto(
                result.Level,
                result.Passed,
                result.ErrorCount,
                result.WarningCount,
                result.Findings.Select(f => new FindingDto(f.Check, f.Severity.ToString(), f.Message, f.Location)).ToList(),
                ToTargetDto(target));

            return ToolResponse.Success(dto);
        });
    }

    /// <summary>
    /// Reads <c>validate_layout</c>'s <c>level</c>, case-insensitively: <paramref name="full"/> is true for
    /// <c>full</c>, false for <c>quick</c>; false for anything else.
    /// </summary>
    internal static bool TryParseLevel(string level, out bool full)
    {
        full = string.Equals(level, "full", StringComparison.OrdinalIgnoreCase);
        return full || string.Equals(level, "quick", StringComparison.OrdinalIgnoreCase);
    }
}
