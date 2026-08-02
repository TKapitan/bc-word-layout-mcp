namespace BcWordLayout.Domain.Models;

/// <summary>
/// Identity of a BC report as parsed from the dataset custom XML part namespace
/// (<c>urn:microsoft-dynamics-nav/reports/&lt;ReportName&gt;/&lt;ReportId&gt;/</c>).
/// </summary>
public sealed class ReportIdentity
{
    public required string ReportName { get; init; }
    public required string ReportId { get; init; }

    /// <summary>The raw dataset namespace URI.</summary>
    public required string Namespace { get; init; }

    /// <summary>
    /// The <c>ds:itemID</c> GUID of the BC custom XML part (from <c>itemPropsN.xml</c>).
    /// Null when the source is a bare exported schema XML (no OPC package, no item props).
    /// </summary>
    public string? StoreItemId { get; init; }
}

/// <summary>A leaf column of a report data item (a bound-able field or a label column).</summary>
public sealed class DatasetColumn
{
    public required string Name { get; init; }

    /// <summary>Slash-delimited path from the root, e.g. <c>/Header/Line/ItemNo_Line</c>.</summary>
    public required string Path { get; init; }

    /// <summary>
    /// True when this column is a label per the active <see cref="LabelConvention.Current"/> — by default a
    /// name ending in <c>Lbl</c>/<c>_Lbl</c>, but see that type for the full (configurable) rule, which also
    /// considers this column's parent data item via <see cref="Path"/>.
    /// </summary>
    public bool IsLabel => LabelConvention.Current.IsLabelPath(Path);
}

/// <summary>
/// A node in the report dataset hierarchy. The root node is the <c>NavWordReportXmlPart</c>;
/// its children are the system node (<c>BCReportInformation</c>, flagged <see cref="IsSystem"/>)
/// and the business data items. Every element with child elements is a data item; leaf elements
/// are surfaced as <see cref="Columns"/>.
/// </summary>
public sealed class DataItem
{
    public required string Name { get; init; }

    /// <summary>Slash-delimited path from the root, e.g. <c>/Header/Line</c>. The root's path is <c>/</c>.</summary>
    public required string Path { get; init; }

    /// <summary>True for the <c>BCReportInformation</c> subtree (system metadata, not business data).</summary>
    public bool IsSystem { get; init; }

    public List<DataItem> Children { get; } = new();

    public List<DatasetColumn> Columns { get; } = new();

    /// <summary>Finds a direct child data item by name (ordinal), or null.</summary>
    public DataItem? FindChildItem(string name) =>
        Children.FirstOrDefault(c => string.Equals(c.Name, name, StringComparison.Ordinal));

    /// <summary>Finds a direct child column by name (ordinal), or null.</summary>
    public DatasetColumn? FindChildColumn(string name) =>
        Columns.FirstOrDefault(c => string.Equals(c.Name, name, StringComparison.Ordinal));
}

/// <summary>A fully parsed report dataset: identity + the data-item hierarchy.</summary>
public sealed class DatasetTree
{
    public required ReportIdentity Report { get; init; }

    /// <summary>The <c>NavWordReportXmlPart</c> root node.</summary>
    public required DataItem Root { get; init; }

    /// <summary>
    /// Enumerates all data items (depth-first). System nodes are included only when
    /// <paramref name="includeSystem"/> is true.
    /// </summary>
    public IEnumerable<DataItem> AllDataItems(bool includeSystem = false)
    {
        return Walk(Root);

        IEnumerable<DataItem> Walk(DataItem node)
        {
            if (node.IsSystem && !includeSystem)
            {
                yield break;
            }

            yield return node;
            foreach (var child in node.Children)
            {
                foreach (var d in Walk(child))
                {
                    yield return d;
                }
            }
        }
    }

    /// <summary>
    /// Enumerates all leaf columns across the tree. The system (<c>BCReportInformation</c>) subtree
    /// is excluded unless <paramref name="includeSystem"/> is true.
    /// </summary>
    public IEnumerable<DatasetColumn> AllColumns(bool includeSystem = false)
    {
        foreach (var item in AllDataItems(includeSystem))
        {
            foreach (var col in item.Columns)
            {
                yield return col;
            }
        }
    }
}
