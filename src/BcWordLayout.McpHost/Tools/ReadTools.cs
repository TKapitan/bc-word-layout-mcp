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
                 + "storeItemID, the full control inventory (kind, alias, tag, xpath, storeItemID, part, "
                 + "parent repeater, and for each control its structural LEVEL - run/block/cell/row/runRuby - "
                 + "plus its tableIndex/rowIndex/colIndex when it sits in a table), a TABLES section "
                 + "describing every table's grid (rowCount, columnCount, gridColumnWidths) and, per row and "
                 + "cell, which control owns it and the cell's visible text, and a quick validation status "
                 + "summary. IMPORTANT for editing: a control whose level is 'cell' wraps a whole table cell "
                 + "(one column) - e.g. BC header address fields; removing it is safe (remove_control never "
                 + "deletes the cell/column), but use this table detail to understand what a control occupies "
                 + "before editing.")]
    public static ToolResponse GetLayoutInfo(
        [Description("Absolute path to the .docx layout file.")] string layoutPath)
    {
        return GuardRead(layoutPath, () =>
        {
            if (!File.Exists(layoutPath))
            {
                throw new FileNotFoundException("layoutPath does not point to an existing file.", layoutPath);
            }

            // Open the package once and share the handle across all three passes (schema, inventory,
            // validation) instead of re-parsing the same .docx repeatedly.
            using var doc = WordprocessingDocument.Open(layoutPath, false);
            var tree = SchemaProvider.FromLayout(doc);
            var inventory = LayoutReader.Read(doc);
            var validation = LayoutValidator.Quick(doc);

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
                Tables: inventory.Tables.Select(ToTableDto).ToList());

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
                "invalid_argument",
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
            using var doc = WordprocessingDocument.Open(source, false);
            var tree = SchemaProvider.FromLayout(doc);
            var boundPaths = BuildBoundPaths(LayoutReader.Read(doc));

            var root = ToDataItemDto(tree.Root, boundPaths);
            var dto = new DatasetFieldsDto("layout", ToReportDto(tree.Report), root);
            return ToolResponse.Success(dto);
        });
    }

    [McpServerTool(Name = "validate_layout")]
    [Description("Validate a BC Word layout. level='quick' runs structural + binding checks (OpenXML "
                 + "validity, single BC XML part, storeItemID match, binding dataset-namespace match, XPath "
                 + "resolution, repeater shape, attachedTemplate warning). level='full' additionally runs a real dry-run merge (sample "
                 + "data fill + repeater expansion) against a throwaway temp copy and surfaces every merge "
                 + "warning (e.g. unresolved bindings, XPath errors) as a validation finding.")]
    public static ToolResponse ValidateLayout(
        [Description("Absolute path to the .docx layout file.")] string layoutPath,
        [Description("Validation level: 'quick' or 'full'. Default 'quick'.")] string level = "quick")
    {
        if (!string.Equals(level, "full", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(level, "quick", StringComparison.OrdinalIgnoreCase))
        {
            return Guard(() => ToolResponse.Failure(
                "invalid_argument",
                $"Unknown validation level '{level}'.",
                "level must be 'quick' or 'full' (case-insensitive); 'quick' runs fast structural+binding "
                + "checks, 'full' additionally dry-run merges sample data and surfaces merge warnings."), layoutPath);
        }

        return GuardRead(layoutPath, () =>
        {
            // FullValidator.Full also reads layoutPath (via a File.Copy dry-run merge) on top of the quick
            // checks it starts with, so the SAME lock pair covers both levels rather than only the quick one.
            var result = string.Equals(level, "full", StringComparison.OrdinalIgnoreCase)
                ? FullValidator.Full(layoutPath)
                : LayoutValidator.Quick(layoutPath);

            var dto = new ValidationResultDto(
                result.Level,
                result.Passed,
                result.ErrorCount,
                result.WarningCount,
                result.Findings.Select(f => new FindingDto(f.Check, f.Severity.ToString(), f.Message, f.Location)).ToList());

            return ToolResponse.Success(dto);
        });
    }
}
