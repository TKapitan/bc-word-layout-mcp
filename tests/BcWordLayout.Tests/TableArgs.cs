using BcWordLayout.Domain.Models;
using BcWordLayout.McpHost.Tools;

namespace BcWordLayout.Tests;

/// <summary>
/// Short forms of the table tools' array arguments (GitHub issue #133) for the cases tests repeat most: a repeater
/// table's columns given by the leaf names of one data item, and a plain table's columns with every default. Any
/// width, alignment or cell is spelled out at the call site, in the shape an agent sends.
/// </summary>
internal static class TableArgs
{
    /// <summary><c>insert_repeater_table</c> columns: one per leaf of <paramref name="dataItem"/>, as full paths.</summary>
    public static List<RepeaterColumnArgument> FieldsOf(string dataItem, params string[] leaves) =>
        [.. leaves.Select(leaf => new RepeaterColumnArgument { Field = $"{dataItem}/{leaf}" })];

    /// <summary>The same columns as the domain takes them, for direct <c>LayoutEditor</c>/<c>SdtFactory</c> calls.</summary>
    public static List<TableColumn> DomainFieldsOf(string dataItem, params string[] leaves) =>
        [.. leaves.Select(leaf => new TableColumn { Field = $"{dataItem}/{leaf}" })];

    /// <summary><c>insert_table</c> columns with every default: an even split, left-aligned.</summary>
    public static List<TableColumnArgument> Columns(int count) =>
        [.. Enumerable.Range(0, count).Select(_ => new TableColumnArgument())];

    /// <summary>The same columns as the domain takes them.</summary>
    public static List<TableColumn> DomainColumns(int count) =>
        [.. Enumerable.Range(0, count).Select(_ => new TableColumn())];
}
