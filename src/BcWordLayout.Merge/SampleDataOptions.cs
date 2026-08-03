using System.Xml.Linq;

namespace BcWordLayout.Merge;

/// <summary>
/// Options controlling <see cref="SampleDataGenerator.Generate"/>. Generation is fully deterministic:
/// the same schema and the same <see cref="Seed"/> always produce byte-identical output.
/// </summary>
public sealed class SampleDataOptions
{
    /// <summary>Seed for the deterministic random generator used to synthesize leaf column values.</summary>
    public int Seed { get; init; } = 12345;

    /// <summary>Number of instances generated per repeating (non-system) business data item.</summary>
    public int Rows { get; init; } = 3;

    /// <summary>
    /// Optional path to a real exported BC dataset XML file to use instead of generating fake data. Both
    /// encodings BC produces are accepted (sniffed by root element): the layout's own data-store part shape
    /// (<c>NavWordReportXmlPart</c>) and the report UI's *Send to → XML* export (<c>ReportDataSet</c>) — the
    /// latter is converted via <see cref="ReportDataSetConverter"/>, which also applies each column's
    /// <c>decimalformatter</c> the way BC's own render does. When set to a non-null, non-empty path,
    /// <see cref="Seed"/> and <see cref="Rows"/> are ignored entirely.
    /// </summary>
    public string? DataOverridesPath { get; init; }

    /// <summary>
    /// Maximum number of instances generated per repeating (non-system) business data item, at ANY level
    /// of the schema tree — a waste-elimination safeguard, normally wired from <see cref="MergeEngine"/>'s
    /// own <see cref="MergeOptions.MaxRowsPerRepeater"/> (same default, 100): that merge-time cap already
    /// discards any row beyond it once cloning into the OOXML tree, so generating more than the merge could
    /// ever use is pure waste — and, for a deeply nested schema, a real risk of building rows^depth
    /// <see cref="System.Xml.Linq.XElement"/> instances before the merge ever gets a chance to trim
    /// anything (a large <see cref="Rows"/> value on a several-levels-deep schema could otherwise exhaust
    /// memory purely generating sample data that would be discarded). Ignored when
    /// <see cref="DataOverridesPath"/> is set (a real exported dataset is loaded verbatim, not generated, so
    /// there is nothing here to cap). Default 100 — comfortably above every corpus/typical <see cref="Rows"/>
    /// value, so an un-configured generation is unaffected.
    /// </summary>
    public int MaxRowsPerItem { get; init; } = 100;

    /// <summary>
    /// A GLOBAL ceiling on the total number of business (non-system) data-item instances generated across the
    /// WHOLE schema tree — the safeguard <see cref="MaxRowsPerItem"/> alone cannot provide. Because generation
    /// creates <c>Math.Min(Rows, MaxRowsPerItem)</c> instances of every child at every nesting level, the
    /// per-item cap bounds each level INDEPENDENTLY but the counts still MULTIPLY across depth
    /// (<c>≈ count^depth</c>): a several-levels-deep schema (e.g. Sales Invoice: Line ▸ Assembly Line ▸
    /// tracking) with a large <see cref="Rows"/> can build hundreds of thousands of
    /// <see cref="System.Xml.Linq.XElement"/> instances — minutes of CPU and hundreds of MB — before the merge
    /// or the per-item cap can trim anything. This budget is checked as instances are created (depth-first, in
    /// document order): once it is exhausted, no further business instances are generated and
    /// <see cref="SampleDataset.Truncated"/> is set so the caller can warn. System subtrees
    /// (<see cref="BcWordLayout.Domain.Models.DataItem.IsSystem"/>, always exactly 1 instance) never count
    /// against it. Ignored when <see cref="DataOverridesPath"/> is set (nothing is generated). Default 20,000 —
    /// far above any real preview need (a genuine BC document rarely exceeds a few hundred rows), yet low enough
    /// that a pathological <c>Rows</c>×depth combination is bounded to well under a second of generation.
    /// </summary>
    public int MaxTotalInstances { get; init; } = 20_000;

    /// <summary>
    /// The set of <see cref="BcWordLayout.Domain.Models.DataItem.Path"/> values the DOCUMENT actually repeats
    /// — i.e. every data item some <c>w15:repeatingSection</c> control's own row binding resolves to,
    /// pre-scanned from the live layout by <see cref="MergeEngine"/> (see
    /// <c>MergeEngine.ScanRepeaterConsumedPaths</c>) before generation runs. This is the waste-elimination
    /// half of the row-multiplication rule: <see cref="SampleDataGenerator.BuildInstance"/> multiplies a business item to
    /// <c>Math.Min(Rows, MaxRowsPerItem)</c> instances ONLY when its own <see cref="BcWordLayout.Domain.Models.DataItem.Path"/>
    /// is in this set; every other business item — one nothing in the document repeats, however deep the
    /// schema nests it — gets exactly ONE instance instead. A single instance is still enough for every other
    /// consumer: a repeater's absolute row xpath always indexes its ancestor steps (e.g. <c>Header[1]</c>), so
    /// an unconsumed ancestor never needed more than one instance in the first place (every extra one was pure
    /// waste, and worse, could exhaust <see cref="MaxTotalInstances"/> before a REAL repeater's own rows were
    /// generated); a field bound under an unconsumed item (no repeater at all in its own subtree) shows a
    /// single resolved value either way; and a completely unread item still gets its one instance, so an
    /// absolute-XPath fallback (<see cref="MergeEngine"/>'s <c>xpath-fallback</c> path, taken when re-anchoring
    /// a divergent binding fails) can still resolve against it instead of false-reporting <c>unresolved-binding</c>.
    /// KNOWN (accepted) limit of the one-instance rule: a binding whose ANCESTOR step carries a positional
    /// index &gt; 1 (e.g. <c>/…/Header[2]/Line</c> as a repeater's own row xpath, or <c>Header[2]/Foo</c> for a
    /// field) would no longer resolve, because the unconsumed ancestor now has exactly one instance — verified
    /// absent from every corpus layout (all 324 bindings index ancestors as <c>[1]</c>), and not something BC
    /// emits; revisit only if a real layout ever surfaces one.
    /// Null (the default — every direct <see cref="SampleDataGenerator.Generate"/> caller that does not scan a
    /// live document, e.g. every existing unit test) means "no scan was done": every business item is treated
    /// as consumed, exactly the pre-B23 behavior, so generation run without a document to scan is unaffected.
    /// </summary>
    public IReadOnlySet<string>? RepeaterConsumedPaths { get; init; }
}

/// <summary>A generated (or loaded) sample dataset, ready for the merge engine to resolve bindings against.</summary>
public sealed class SampleDataset
{
    /// <summary>The dataset XML, rooted at the schema's root element name, in <see cref="Namespace"/>.</summary>
    public required XDocument Xml { get; init; }

    /// <summary>The dataset namespace URI (matches the source schema's <c>Report.Namespace</c>).</summary>
    public required string Namespace { get; init; }

    /// <summary>
    /// True when generation stopped early because <see cref="SampleDataOptions.MaxTotalInstances"/> was
    /// reached — the generated data (and therefore any preview built from it) is deliberately PARTIAL. Always
    /// false for loaded <see cref="SampleDataOptions.DataOverridesPath"/> data (loaded in full, never
    /// generated) and for any generation that stayed within the budget. <see cref="MergeEngine"/> surfaces this
    /// as a <c>sample-data-capped</c> <see cref="MergeWarning"/> so a capped preview is never silent.
    /// </summary>
    public bool Truncated { get; init; }
}
