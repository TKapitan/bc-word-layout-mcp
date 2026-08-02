namespace BcWordLayout.McpHost;

/// <summary>
/// Host-level environment-variable configuration read once at process startup (see <c>Program.cs</c>).
/// Kept as a small, pure, side-effect-free static method (no I/O, no logging) so <c>BcWordLayout.Tests</c>
/// can drive the parsing logic directly — via this project's own <c>InternalsVisibleTo</c> grant to
/// <c>BcWordLayout.Tests</c> — without spawning the actual MCP host process.
/// </summary>
internal static class EnvironmentConfig
{
    /// <summary>
    /// Comma-separated list of label-name suffixes overriding the default BC label convention (a single
    /// suffix, <c>"Lbl"</c>), e.g. <c>"Lbl,Caption"</c>.
    /// Read once at host startup by <c>Program.cs</c> to install a custom
    /// <see cref="BcWordLayout.Domain.LabelConvention"/> on
    /// <see cref="BcWordLayout.Domain.LabelConvention.Current"/>. Configure it per-MCP-server in the
    /// client's MCP config (e.g. the server's <c>env</c> block) — see the README's configuration section.
    /// </summary>
    public const string LabelSuffixesVariable = "BCWL_LABEL_SUFFIXES";

    /// <summary>
    /// Optional name of a dedicated labels DATA ITEM (e.g. <c>"Labels"</c>): every direct column of a data
    /// item with this exact name is classified as a label regardless of its own suffix — the shape the
    /// corpus's <c>InventoryOrderDetails.docx</c> proves is real (its <c>&lt;Labels&gt;</c> item mixes
    /// <c>*Label</c>/<c>*Caption</c>/unsuffixed column names a suffix list alone can never classify). Unset
    /// (the default) disables the rule. Combines freely with <see cref="LabelSuffixesVariable"/>; either may
    /// be set without the other.
    /// </summary>
    public const string LabelsDataItemVariable = "BCWL_LABELS_DATA_ITEM";

    /// <summary>
    /// Parses <paramref name="rawValue"/> (the raw <see cref="LabelSuffixesVariable"/> value) into an
    /// ordered, de-duplicated, trimmed suffix list. Returns null for a null/blank/whitespace-only value, or
    /// one that reduces to nothing usable after trimming/de-duplication (e.g. <c>","</c> or <c>" , "</c>) —
    /// callers must treat null as "keep the current/default convention". Never throws.
    /// </summary>
    public static IReadOnlyList<string>? ParseLabelSuffixes(string? rawValue)
    {
        if (string.IsNullOrWhiteSpace(rawValue))
        {
            return null;
        }

        var suffixes = rawValue
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Distinct(StringComparer.Ordinal)
            .ToArray();

        return suffixes.Length == 0 ? null : suffixes;
    }

    /// <summary>
    /// Parses <paramref name="rawValue"/> (the raw <see cref="LabelsDataItemVariable"/> value) into a
    /// trimmed data-item name. Returns null for a null/blank value or one that is not a plausible single
    /// element name (contains <c>'/'</c> — the rule matches ONE data-item name, never a path) — callers must
    /// treat null as "rule stays disabled". Never throws.
    /// </summary>
    public static string? ParseLabelsDataItemName(string? rawValue)
    {
        if (string.IsNullOrWhiteSpace(rawValue))
        {
            return null;
        }

        var name = rawValue.Trim();
        return name.Contains('/') ? null : name;
    }
}
