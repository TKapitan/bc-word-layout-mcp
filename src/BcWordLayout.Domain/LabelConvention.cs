namespace BcWordLayout.Domain;

/// <summary>
/// Configurable BC label naming convention: the single source of truth every label/field classifier in this
/// codebase reads (via <see cref="Current"/>) to decide whether a dataset column is a LABEL or a plain
/// FIELD.
/// </summary>
/// <remarks>
/// <para>
/// Business Central's own convention is a leaf column name suffix — <c>*_Lbl</c> or <c>*Lbl</c> (a single
/// "ends with Lbl" check covers both, since <c>_Lbl</c> itself ends in <c>Lbl</c>). That is this type's
/// DEFAULT (<see cref="Default"/>): a single suffix, <c>"Lbl"</c>.
/// </para>
/// <para>
/// Real corpus evidence proves BC layouts can name their
/// label columns entirely differently: <c>InventoryOrderDetails.docx</c>'s dataset custom XML carries a
/// dedicated <c>&lt;Labels&gt;</c> data item directly under the root whose ~27 direct columns are ALL
/// captions/labels regardless of their OWN suffix — a mix of <c>*Label</c> (<c>CompanyLabel</c>),
/// <c>*Caption</c> (<c>No_ItemCaption</c>), and even unsuffixed names (<c>DataRetrieved</c>). Under the
/// plain "Lbl" suffix rule alone, none of those are recognized as labels (see the
/// <c>InventoryOrderDetails</c> cases in <c>SchemaProviderTests</c>/<c>LayoutReaderTests</c>, which pin
/// exactly that gap). This type supports BOTH shapes: an ordered <see cref="Suffixes"/> list (checked by
/// <see cref="IsLabelName"/>), and an optional <see cref="LabelsDataItemName"/> rule (checked by
/// <see cref="IsLabelPath(System.Collections.Generic.IReadOnlyList{string})"/>) that treats every DIRECT
/// column of a data item with that exact name as a label unconditionally.
/// </para>
/// <para>
/// <see cref="LabelsDataItemName"/> defaults to <c>null</c> (disabled) even though the corpus proves the
/// shape is real: turning it on BY DEFAULT would flip <c>InventoryOrderDetails.docx</c>'s label-count
/// assertions in <c>SchemaProviderTests</c>/<c>LayoutReaderTests</c> from 0 to ~27 — an unrequested default-
/// behavior change - the existing suite must stay green UNCHANGED. Callers who
/// know their own corpus uses this shape opt in explicitly, e.g.
/// <c>new LabelConvention(suffixes, labelsDataItemName: "Labels")</c>.
/// </para>
/// </remarks>
public sealed class LabelConvention
{
    /// <summary>The BC-standard convention every consumer uses unless <see cref="Current"/> is reassigned.</summary>
    public static readonly LabelConvention Default = new(new[] { "Lbl" });

    private static LabelConvention _current = Default;

    /// <summary>
    /// The process-wide active convention: every classifier in this codebase (<see cref="SdtInspector"/>,
    /// <see cref="SdtFactory"/>, <c>DatasetColumn.IsLabel</c>, <c>BcWordLayout.Merge</c>'s sample-data
    /// generation) reads THIS, never <see cref="Default"/> directly — the same minimal-seam pattern as
    /// <c>LifecycleTools.SelectConverter</c>: a plain settable static, no DI. Defaults to
    /// <see cref="Default"/>. The MCP host's <c>Program.cs</c> reassigns it at startup from the
    /// <c>BCWL_LABEL_SUFFIXES</c> environment variable when present (see <c>EnvironmentConfig</c> in
    /// <c>BcWordLayout.McpHost</c>).
    /// </summary>
    /// <remarks>
    /// Public rather than internal specifically so the host project (a separate assembly with no
    /// <c>InternalsVisibleTo</c> grant from this one) can set it at startup, and so library consumers
    /// outside this repo can configure it too. Tests that reassign it for a scenario MUST restore it in a
    /// <c>finally</c> and join the <c>label-convention-seam</c> xUnit collection alongside every test class
    /// whose assertions are sensitive to the result (currently: any test touching
    /// <c>InventoryOrderDetails.docx</c>, the corpus's only non-"Lbl"-suffixed label-shaped file) — this is
    /// a process-wide static, so an un-restored swap or a racing parallel test class would leak into an
    /// unrelated test's classification.
    /// </remarks>
    public static LabelConvention Current
    {
        get => _current;
        set => _current = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>
    /// Ordered, case-sensitive (ordinal) list of recognized label-name suffixes — a column name ending in
    /// ANY of these is a label (see <see cref="IsLabelName"/>). At least one entry is required.
    /// </summary>
    public IReadOnlyList<string> Suffixes { get; }

    /// <summary>
    /// When set, every DIRECT column of a data item with this exact name (ordinal) is a label regardless of
    /// its own suffix — the shape <c>InventoryOrderDetails.docx</c>'s <c>&lt;Labels&gt;</c> data item proves
    /// is real (see this type's remarks). <c>null</c> disables the rule (suffix-list-only classification,
    /// the BC-standard behavior).
    /// </summary>
    public string? LabelsDataItemName { get; }

    public LabelConvention(IReadOnlyList<string> suffixes, string? labelsDataItemName = null)
    {
        ArgumentNullException.ThrowIfNull(suffixes);
        if (suffixes.Count == 0)
        {
            throw new ArgumentException("At least one label suffix is required.", nameof(suffixes));
        }

        if (suffixes.Any(string.IsNullOrEmpty))
        {
            throw new ArgumentException("Label suffixes must not be null or empty.", nameof(suffixes));
        }

        Suffixes = suffixes.ToArray();
        LabelsDataItemName = string.IsNullOrWhiteSpace(labelsDataItemName) ? null : labelsDataItemName;
    }

    /// <summary>
    /// Returns true when <paramref name="columnName"/> ends with any configured <see cref="Suffixes"/>
    /// entry (ordinal). Does not know about <see cref="LabelsDataItemName"/> — see
    /// <see cref="IsLabelPath(System.Collections.Generic.IReadOnlyList{string})"/> for the full rule when
    /// the column's data-item context is available.
    /// </summary>
    public bool IsLabelName(string? columnName) =>
        !string.IsNullOrEmpty(columnName)
        && Suffixes.Any(suffix => columnName.EndsWith(suffix, StringComparison.Ordinal));

    /// <summary>
    /// Full label classification for a column given its ordered path segments (leaf last; parent data items
    /// before it — the root name may or may not be included, only the IMMEDIATE parent, <c>segments[^2]</c>,
    /// matters). A column is a label when either its own name matches <see cref="IsLabelName"/>, OR its
    /// direct parent data item is named <see cref="LabelsDataItemName"/> (when that rule is enabled).
    /// </summary>
    public bool IsLabelPath(IReadOnlyList<string> segments)
    {
        if (segments is null || segments.Count == 0)
        {
            return false;
        }

        if (segments.Count >= 2 && LabelsDataItemName is not null
            && string.Equals(segments[^2], LabelsDataItemName, StringComparison.Ordinal))
        {
            return true;
        }

        return IsLabelName(segments[^1]);
    }

    /// <summary>
    /// Convenience overload for a slash-delimited path string (e.g. <c>DatasetColumn.Path</c>, a binding
    /// xpath already reduced to plain segment names).
    /// </summary>
    public bool IsLabelPath(string? path) =>
        IsLabelPath(string.IsNullOrEmpty(path)
            ? Array.Empty<string>()
            : path.Split('/', StringSplitOptions.RemoveEmptyEntries));

    /// <summary>
    /// Strips the first configured suffix (in <see cref="Suffixes"/> order) that <paramref name="name"/>
    /// ends with, plus a joining underscore if one remains (so <c>_Lbl</c> and bare <c>Lbl</c> both collapse
    /// to the same base name); returns <paramref name="name"/> unchanged if no suffix matches — e.g. a label
    /// recognized only via <see cref="LabelsDataItemName"/> with no matching suffix at all (like
    /// <c>DataRetrieved</c>).
    /// </summary>
    public string StripSuffix(string name)
    {
        ArgumentNullException.ThrowIfNull(name);
        foreach (var suffix in Suffixes)
        {
            if (name.Length > suffix.Length && name.EndsWith(suffix, StringComparison.Ordinal))
            {
                var trimmed = name[..^suffix.Length];
                return trimmed.EndsWith('_') ? trimmed[..^1] : trimmed;
            }
        }

        return name;
    }

    /// <summary>
    /// Human-readable description of this convention's rule, for dynamic error/hint text (see
    /// <c>ToolGuards.InvalidArgumentHint</c> and <c>SdtFactory</c>'s label/field <c>ArgumentException</c>s)
    /// so those messages never hard-promise "Lbl" once a custom convention is installed.
    /// </summary>
    public string Describe()
    {
        var suffixPart = "a name ending in " + string.Join(" or ", Suffixes.Select(s => $"'{s}'"));
        return LabelsDataItemName is null
            ? suffixPart
            : $"{suffixPart}, or any column directly under a '{LabelsDataItemName}' data item";
    }
}
