using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;

namespace BcWordLayout.McpHost.Tools;

/// <summary>
/// One column of <c>insert_repeater_table</c>'s <c>columns</c> array (GitHub issue #133): the column's field and,
/// optionally, its width and alignment, together, instead of three parallel comma-separated lists that had to be
/// kept in step by hand. Its member descriptions are what an agent reads in the tool schema.
/// </summary>
public sealed record RepeaterColumnArgument
{
    [Description("Full dataset path of a leaf column of dataItem, exactly as list_dataset_fields reports it, e.g. "
                 + "'/Header/Line/ItemNo_Line'. The header cell binds its label column when one exists.")]
    public required string Field { get; init; }

    [Description(TableArguments.WidthDescription)]
    [Range(1, TableTools.MaxColumnWidthTwips)]
    public int? Width { get; init; }

    [Description("'left', 'center' or 'right', for the header cell and the data cell alike - real BC line tables "
                 + "right-align their quantity, price and amount columns. Omit for left.")]
    [AllowedValues(TableArguments.Left, TableArguments.Center, TableArguments.Right, null)]
    public string? Alignment { get; init; }
}

/// <summary>One column of <c>insert_table</c>'s <c>columns</c> array: its width and alignment, both optional.</summary>
public sealed record TableColumnArgument
{
    [Description(TableArguments.WidthDescription)]
    [Range(1, TableTools.MaxColumnWidthTwips)]
    public int? Width { get; init; }

    [Description("'left', 'center' or 'right' for every cell of the column; set_cell_text and insert_field keep it. "
                 + "Omit for left.")]
    [AllowedValues(TableArguments.Left, TableArguments.Center, TableArguments.Right, null)]
    public string? Alignment { get; init; }
}

/// <summary>
/// One cell of a new row's <c>cells</c> array (<c>insert_repeater_row</c>, <c>insert_subtotal_row</c>,
/// <c>insert_table_row</c>): the fields it binds, how many grid columns it spans and its alignment. It replaces the
/// <c>'-'</c>, <c>'a+b'</c> and <c>'N:'</c> mini-grammar the comma-separated string carried.
/// </summary>
public sealed record RowCellArgument
{
    [Description("Full dataset paths bound in this cell, in order, exactly as list_dataset_fields reports them; "
                 + "several are chained inline with no separator (the corpus 'caption: value' shape), and a "
                 + "label-shaped path becomes a label control. Omit, or pass [], for an empty spacer cell.")]
    public IReadOnlyList<string>? Fields { get; init; }

    [Description("How many grid columns the cell spans, 1 to 63; default 1. The row's spans must add up to the "
                 + "table's columnCount.")]
    [Range(1, TableTools.MaxCellSpan)]
    public int? Span { get; init; }

    [Description("'left', 'center' or 'right'; omit for the default (left). Amounts are right-aligned in the corpus.")]
    [AllowedValues(TableArguments.Left, TableArguments.Center, TableArguments.Right, null)]
    public string? Alignment { get; init; }
}

/// <summary>
/// Shared descriptions, values and conversions for the table tools' array arguments. The conversions refuse what
/// the binder lets through — a missing or empty array, a <c>null</c> entry, a number out of range — naming the
/// entry (<c>cells[2].span</c>); the domain checks the rest (fields that resolve, spans that cover the grid,
/// alignments) and names entries the same way.
/// </summary>
internal static class TableArguments
{
    internal const string Left = "left";
    internal const string Center = "center";
    internal const string Right = "right";

    internal const string WidthDescription =
        "Width in twips (1/20 pt), 1 to 31680. Omit for an even share of the width the other columns leave. The "
        + "widths are binding: the table is emitted with w:tblLayout='fixed', so Word does not autofit them.";

    /// <summary><c>insert_repeater_table</c>'s <c>columns</c> as the domain's per-column specs.</summary>
    internal static List<TableColumn> ToRepeaterColumns(IReadOnlyList<RepeaterColumnArgument?>? columns) =>
        Entries(columns, "columns", "e.g. [{\"field\": \"/Header/Line/ItemNo_Line\"}]")
            .Select((column, i) => new TableColumn
            {
                Field = column.Field,
                Width = Width(column.Width, i),
                Alignment = column.Alignment,
            })
            .ToList();

    /// <summary><c>insert_table</c>'s <c>columns</c> as the domain's per-column specs.</summary>
    internal static List<TableColumn> ToTableColumns(IReadOnlyList<TableColumnArgument?>? columns) =>
        Entries(columns, "columns", "e.g. [{}, {\"width\": 2400, \"alignment\": \"right\"}]")
            .Select((column, i) => new TableColumn { Width = Width(column.Width, i), Alignment = column.Alignment })
            .ToList();

    /// <summary>The row tools' <c>cells</c> as the domain's per-cell specs.</summary>
    internal static List<RepeaterRowCell> ToRowCells(IReadOnlyList<RowCellArgument?>? cells) =>
        Entries(cells, "cells", "e.g. [{\"span\": 3}, {\"fields\": [\"/Header/Totals/TotalAmount\"]}]")
            .Select((cell, i) => new RepeaterRowCell
            {
                Span = Span(cell.Span, i, nameof(cells)),
                Fields = Fields(cell.Fields, i, nameof(cells)),
                Alignment = cell.Alignment,
            })
            .ToList();

    private static List<T> Entries<T>(IReadOnlyList<T?>? entries, string name, string example)
        where T : class
    {
        if (entries is null || entries.Count == 0)
        {
            throw new ArgumentException($"{name} must hold at least one entry, {example}.", name);
        }

        for (var i = 0; i < entries.Count; i++)
        {
            if (entries[i] is null)
            {
                throw new ArgumentException(
                    $"{name}[{i}] is null; describe every entry as an object, {example} ({{}} takes every default).",
                    name);
            }
        }

        return entries.Select(e => e!).ToList();
    }

    private static int? Width(int? width, int index) =>
        width is { } twips ? TableTools.CheckColumnWidth(twips, $"columns[{index}].width ({twips})", "columns") : null;

    private static int Span(int? span, int index, string paramName) => span switch
    {
        null => 1,
        >= 1 and <= TableTools.MaxCellSpan => span.Value,
        _ => throw new ArgumentException(
            $"cells[{index}].span is {span}; a cell spans between 1 and {TableTools.MaxCellSpan} grid columns (omit "
            + "span for one).",
            paramName),
    };

    private static List<string> Fields(IReadOnlyList<string?>? fields, int index, string paramName)
    {
        if (fields is null)
        {
            return [];
        }

        for (var f = 0; f < fields.Count; f++)
        {
            if (string.IsNullOrWhiteSpace(fields[f]))
            {
                throw new ArgumentException(
                    $"cells[{index}].fields[{f}] is empty; each entry is the full dataset path of a field, as "
                    + "list_dataset_fields reports it (omit fields, or pass [], for an empty spacer cell).",
                    paramName);
            }
        }

        return fields.Select(f => f!).ToList();
    }
}
