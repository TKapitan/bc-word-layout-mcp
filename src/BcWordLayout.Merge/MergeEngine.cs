using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Blip = DocumentFormat.OpenXml.Drawing.Blip;
using DrawWordprocessing = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace BcWordLayout.Merge;

/// <summary>Options controlling <see cref="MergeEngine.Merge(string, string, MergeOptions?)"/>.</summary>
public sealed class MergeOptions
{
    /// <summary>Seed for the deterministic sample-data generator (see <see cref="SampleDataOptions.Seed"/>).</summary>
    public int Seed { get; init; } = 12345;

    /// <summary>Number of instances generated per repeating business data item.</summary>
    public int Rows { get; init; } = 3;

    /// <summary>Optional real exported dataset XML to merge with instead of generated sample data.</summary>
    public string? DataOverridesPath { get; init; }

    /// <summary>
    /// Maximum number of row clones the merge will generate for any SINGLE repeating section — applied
    /// independently to EACH repeater encountered (including every nested repeater's own row clones, once
    /// per surviving enclosing row — each such expansion counts as its own entry in
    /// <see cref="MergeStats.RepeatersExpanded"/>), not as one global total across the whole document. A
    /// deeply-nested layout's row count multiplies across nesting depth (rows^depth), so a large
    /// <see cref="Rows"/> value (or, with <see cref="DataOverridesPath"/>, a real exported dataset that
    /// simply happens to contain an unusually large number of matching rows) can otherwise blow up preview
    /// time/memory. When a repeater would otherwise generate more rows than this, only the first
    /// <see cref="MaxRowsPerRepeater"/> are cloned and a <c>row-cap</c> <see cref="MergeWarning"/> naming
    /// the repeater and the cap is always raised — capping never happens silently. Default 100, comfortably
    /// above any real BC layout's typical preview row count.
    /// </summary>
    public int MaxRowsPerRepeater { get; init; } = 100;

    /// <summary>
    /// Global ceiling on the TOTAL number of generated business (non-system) data-item instances across the
    /// whole schema — wired straight into <see cref="SampleDataOptions.MaxTotalInstances"/>. Unlike
    /// <see cref="MaxRowsPerRepeater"/> (a PER-item cap), this bounds the <c>count^depth</c> multiplicative
    /// blow-up a deeply-nested layout would otherwise suffer for a large <see cref="Rows"/> value, keeping
    /// sample-data generation (and therefore preview time/memory) bounded regardless of nesting depth. Ignored
    /// for <see cref="DataOverridesPath"/> (nothing is generated). When it trims generation, a
    /// <c>sample-data-capped</c> <see cref="MergeWarning"/> is raised — capping is never silent. Default 20,000.
    /// </summary>
    public int MaxTotalInstances { get; init; } = 20_000;

    /// <summary>
    /// When true, after filling controls and cloning repeater rows the merge SEVERS every live data binding
    /// (<c>w:dataBinding</c> / <c>w15:dataBinding</c>) and repeating-section marker
    /// (<c>w15:repeatingSection</c> / <c>w15:repeatingSectionItem</c>) from the output, producing a
    /// self-contained STATIC snapshot. This is required for a rendered preview: BC layout controls are live
    /// data-bound, so when Word (or another renderer) opens the merged file it re-syncs every control from the
    /// mapped custom XML data part and re-evaluates each repeating section against it — and since the merge
    /// writes VISIBLE run text rather than repopulating that part, an un-flattened file would render the
    /// field-name placeholders and a single template row instead of the merged sample data and its rows.
    /// Defaults to <c>false</c> so the logical merge output (with bindings intact) is preserved for validation
    /// and structural inspection; <c>preview_layout</c> sets it <c>true</c>.
    /// </summary>
    public bool FlattenBindingsForRender { get; init; }

    /// <summary>
    /// When true, removes every external relationship (<c>attachedTemplate</c>, externally-linked images,
    /// linked OLE objects, frame/subDocument references — see <see cref="ExternalRelationshipStripper"/> for
    /// the full strip/keep rationale) from the merged output before the merge returns, so a renderer that
    /// opens the merged copy afterward can never dereference one: a poisoned layout's
    /// <c>attachedTemplate</c> pointing at a UNC path or URL would otherwise make Word reach out when it
    /// opens the file, leaking the developer's NTLM hash over SMB or acting as SSRF. Word's
    /// <c>AutomationSecurity</c> does not help — it blocks macros, not template/linked-resource loading. Plain <c>hyperlink</c>
    /// relationships are never affected either way - see the stripper's own remarks. Each relationship
    /// removed raises an <c>external-relationship-stripped</c> <see cref="MergeWarning"/> so the caller knows
    /// the preview differs from the source layout in this respect. Defaults to <c>false</c> so
    /// <c>validate_layout level=full</c>'s dry-run merge (<c>BcWordLayout.Merge.FullValidator</c>) is
    /// unaffected - its throwaway output is deleted in a <c>finally</c> block and never opened by any
    /// converter, so stripping there would add cost with no security benefit; <c>preview_layout</c> sets this
    /// <c>true</c> because ITS merged copy IS opened by a real converter.
    /// </summary>
    public bool StripExternalRelationships { get; init; }
}

/// <summary>A non-fatal finding raised while merging (e.g. a binding that did not resolve).</summary>
public sealed class MergeWarning
{
    /// <summary>Short machine-readable category, e.g. <c>"unresolved-binding"</c>.</summary>
    public required string Kind { get; init; }

    /// <summary>Human-readable description of the finding.</summary>
    public required string Message { get; init; }

    /// <summary>Where the finding occurred, formatted as <c>"part: xpath"</c>, or null.</summary>
    public string? Location { get; init; }
}

/// <summary>Counters summarizing what a merge did.</summary>
public sealed class MergeStats
{
    /// <summary>Bound field/label controls whose XPath resolved and whose text was set.</summary>
    public int FieldsFilled { get; init; }

    /// <summary>Repeating-section controls expanded (each nested clone counts as its own expansion).</summary>
    public int RepeatersExpanded { get; init; }

    /// <summary>Total data rows produced across every repeater expansion.</summary>
    public int RowsGenerated { get; init; }

    /// <summary>Bound controls whose XPath did not resolve against the sample dataset.</summary>
    public int Unresolved { get; init; }

    /// <summary>Picture controls whose blip was repointed to the embedded placeholder image.</summary>
    public int PicturesFilled { get; init; }
}

/// <summary>The outcome of a merge: summary counters plus every non-fatal finding.</summary>
public sealed class MergeResult
{
    public required MergeStats Stats { get; init; }

    public required IReadOnlyList<MergeWarning> Warnings { get; init; }
}

/// <summary>
/// Fills a BC Word layout with sample data by walking the live OOXML tree of the main document plus every
/// header and footer: bound field/label controls get their visible text set, and repeating sections are
/// expanded row by row by cloning their row template once per selected data node. Inner bindings inside a
/// repeater store fully absolute XPaths, so each clone's bindings are re-anchored to that row's own data
/// node rather than the document root — see <see cref="XPathReanchor"/> — which is what makes arbitrary
/// nesting depth (the corpus reaches 3–4 levels) resolve correctly. Picture controls have their blip
/// repointed to a small embedded placeholder image (<see cref="PlaceholderImage"/>), one placeholder
/// <see cref="ImagePart"/> added and cached per hosting part and reused by every picture in it, so a
/// downstream converter renders a visible box instead of BC's real (10-byte stub) image reference.
/// </summary>
public static class MergeEngine
{
    private static readonly Regex PrefixMappingPattern =
        new(@"xmlns:(?<prefix>\w+)\s*=\s*'(?<uri>[^']*)'", RegexOptions.Compiled);

    /// <summary>
    /// Copies <paramref name="layoutPath"/>, fills the copy with a generated (or overridden) sample dataset,
    /// and — only once that fully succeeds — places the merged result at <paramref name="outputPath"/>.
    /// </summary>
    /// <remarks>
    /// <para>
    /// ATOMICITY. The fill happens against a throwaway staging copy in
    /// <paramref name="outputPath"/>'s own directory, never against <paramref name="outputPath"/> itself; the
    /// staging file is moved onto <paramref name="outputPath"/> as the last step, after every part is saved,
    /// and is always deleted (whether that move happened or not) before this method returns or throws. A
    /// mid-merge exception (e.g. <see cref="LayoutValidator"/>/<see cref="ResourceLimits"/> depth guards)
    /// therefore leaves <paramref name="outputPath"/> exactly as it was before the call — absent if it did
    /// not already exist, unchanged if it did — never a partially-merged file.
    /// </para>
    /// <para>
    /// Bounds generation itself, not just cloning: <see cref="SampleDataOptions.MaxRowsPerItem"/> is wired
    /// from <paramref name="options"/>'s own <see cref="MergeOptions.MaxRowsPerRepeater"/>, so
    /// <see cref="SampleDataGenerator"/> never builds more than that many instances of any repeating item
    /// in the first place — waste elimination, since <see cref="Merge(WordprocessingDocument, SampleDataset, MergeOptions?)"/>'s
    /// own per-repeater cap would only have discarded the excess at clone time anyway (this changes no
    /// OUTPUT: the final merged row count is still <c>Math.Min(Rows, MaxRowsPerRepeater)</c> either way).
    /// One consequence: for GENERATED data (no <see cref="MergeOptions.DataOverridesPath"/>), the per-
    /// repeater clone-time check can no longer itself observe more matched rows than the cap (generation
    /// already limited them), so it can no longer raise its own <c>row-cap</c> warning for this path. To
    /// still honor "no silent caps" end to end, THIS method raises one document-level <c>row-cap</c> warning
    /// up front whenever <see cref="MergeOptions.Rows"/> itself exceeds <see cref="MergeOptions.MaxRowsPerRepeater"/>
    /// and no overrides are in play — see <see cref="AppendRequestedRowsExceedsCapWarningIfNeeded"/>. This
    /// does not apply to <see cref="MergeOptions.DataOverridesPath"/>: a real exported dataset is loaded
    /// verbatim (never capped at generation, since there is no generation), so the existing per-repeater
    /// clone-time check remains the live, necessary safeguard — and its own warning — for that path.
    /// </para>
    /// </remarks>
    public static MergeResult Merge(string layoutPath, string outputPath, MergeOptions? options = null)
    {
        if (!File.Exists(layoutPath))
        {
            throw new FileNotFoundException("Layout file not found.", layoutPath);
        }

        if (string.Equals(Path.GetFullPath(layoutPath), Path.GetFullPath(outputPath), StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException("outputPath must differ from layoutPath.", nameof(outputPath));
        }

        options ??= new MergeOptions();

        // Merge into a STAGING copy, same directory as outputPath (guaranteeing a same-volume,
        // effectively-atomic replace), never into outputPath directly. The staging document keeps its
        // default AutoSave=true — deliberately NOT flipped off here, because the explicit per-part Save()
        // calls below coexist with parts the ExternalRelationshipStripper/id-regeneration may touch
        // by no name this method holds (see the DocumentSettingsPart remark below) — AutoSave is the
        // documented backstop for exactly those, and disabling it would silently drop such an edit on a
        // SUCCESSFUL merge. Instead, atomicity comes from staging: if anything throws mid-merge, AutoSave
        // may still flush a half-merged tree into the STAGING file on Dispose, but that file is deleted in
        // the `finally` below and outputPath — pre-existing or not — is never touched. Only a clean, fully
        // completed merge is copied over outputPath, exactly once, at the very end.
        var outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));

        // Best-effort orphan cleanup (interaction): a hard kill between the
        // File.Copy and File.Move below leaves THIS method's own `.bcwl-merge-stage-*.docx` behind, same
        // failure mode ToolGuards' `.bcwl-stage-*.docx` sweep exists for — see SweepStaleMergeStagingFiles.
        if (!string.IsNullOrEmpty(outputDir))
        {
            SweepStaleMergeStagingFiles(outputDir);
        }

        var stagingPath = string.IsNullOrEmpty(outputDir)
            ? $".bcwl-merge-stage-{Guid.NewGuid():N}.docx"
            : Path.Combine(outputDir, $".bcwl-merge-stage-{Guid.NewGuid():N}.docx");

        MergeResult result;
        try
        {
            File.Copy(layoutPath, stagingPath, overwrite: true);
            result = MergeIntoOpenCopy(stagingPath, options);

            // Only reached once every mutation/save above completed without throwing: the staged file is a
            // fully-merged, self-consistent document. Replace outputPath with it as the last step — a plain
            // File.Move(overwrite:true) between two paths in the same directory, i.e. the same volume.
            File.Move(stagingPath, outputPath, overwrite: true);
        }
        finally
        {
            if (File.Exists(stagingPath))
            {
                File.Delete(stagingPath);
            }
        }

        return result;
    }

    /// <summary>
    /// Retention window for <see cref="SweepStaleMergeStagingFiles"/> — see that method's own remarks. Same
    /// value and rationale as <c>BcWordLayout.McpHost.Tools.ToolGuards.StageFileRetentionWindow</c> (the
    /// projects cannot share the constant directly, but both sweeps guard the identical failure mode: a
    /// crash between a stage-copy and its atomic rename).
    /// </summary>
    /// <remarks>
    /// AGE SIGNAL — <see cref="File.GetCreationTimeUtc(string)"/>, deliberately NOT
    /// <see cref="File.GetLastWriteTimeUtc(string)"/> (a bug found and fixed during review of this exact
    /// sweep): <see cref="File.Copy(string, string, bool)"/> copies the SOURCE's last-write time onto the
    /// new file but resets its creation time to "now" (confirmed empirically on Windows/NTFS). The staged
    /// file this sweep targets is itself produced by copying <c>layoutPath</c> — if that source is an old,
    /// already-committed file (the common case: a real corpus/checked-in layout), its last-write time can
    /// trivially be more than a day old, making a staging file created THIS INSTANT already look "stale" by
    /// last-write time alone — exactly the failure mode that let a concurrent sweep delete a live, mid-merge
    /// staging file out from under an in-progress merge. Creation time reliably reflects how long THIS
    /// staging file has actually existed, regardless of its content's own history.
    /// </remarks>
    private static readonly TimeSpan MergeStagingRetentionWindow = TimeSpan.FromDays(1);

    /// <summary>
    /// Best-effort age-based cleanup of orphaned <c>.bcwl-merge-stage-*.docx</c> files (Opus review of B21,
    /// <see cref="Merge(string, string, MergeOptions?)"/> stages its fill into
    /// exactly this shape of file, in <paramref name="outputDir"/>, and normally deletes it again within the
    /// same call (consumed by the commit <see cref="File.Move(string, string, bool)"/> on success, or the
    /// <c>finally</c> block's explicit delete on failure). A process kill in the narrow window between the
    /// stage-copy and either of those leaves the file behind — next to a preview/merge output directory that
    /// may itself sit inside a source-controlled workspace, where it never self-heals and shows up in
    /// <c>git status</c>. This sweep runs on EVERY file-path merge (before that call creates its own staged
    /// file), scoped ONLY to <paramref name="outputDir"/> — the directory this call's own output is about to
    /// land in anyway, never a layout's directory or a caller's wider workspace (that is a DIFFERENT
    /// directory in the common case, e.g. <c>preview_layout</c>'s default output root; see
    /// <c>ToolGuards.SweepStaleStagingFiles</c>, which additionally sweeps this SAME filename shape from a
    /// layout's own directory as a complementary, opportunistic pass — for the narrower case where a caller
    /// happens to point a merge/preview <c>outputDir</c> at the layout's own directory).
    /// </summary>
    /// <remarks>
    /// <para>
    /// SAFE AGAINST A LIVE CONCURRENT MERGE: only files older than <see cref="MergeStagingRetentionWindow"/>
    /// are removed — comfortably longer than any plausible in-flight merge — so a staging file another
    /// call (this process or a different one, merging some OTHER layout into an output that happens to
    /// share this directory) is actively producing can never be mistaken for an orphan.
    /// </para>
    /// <para>
    /// BEST-EFFORT: every failure — a locked file, an inaccessible directory, a candidate removed by a
    /// concurrently-running sweep — is swallowed. A missed sweep simply leaves the file for a later merge
    /// call to retry; it must never fail the merge this cleanup piggybacks on.
    /// </para>
    /// </remarks>
    private static void SweepStaleMergeStagingFiles(string outputDir)
    {
        try
        {
            if (!Directory.Exists(outputDir))
            {
                return;
            }

            var cutoffUtc = DateTime.UtcNow - MergeStagingRetentionWindow;
            foreach (var candidate in Directory.EnumerateFiles(outputDir, ".bcwl-merge-stage-*.docx"))
            {
                try
                {
                    // CreationTimeUtc, not LastWriteTimeUtc — see MergeStagingRetentionWindow's remarks.
                    if (File.GetCreationTimeUtc(candidate) < cutoffUtc)
                    {
                        File.Delete(candidate);
                    }
                }
                catch
                {
                    // Best-effort: leave this one candidate for a later sweep rather than failing the merge.
                }
            }
        }
        catch
        {
            // Best-effort: an inaccessible/unreadable directory must never fail the merge.
        }
    }

    /// <summary>
    /// Opens <paramref name="stagingPath"/> (already a fresh copy of the source layout — see
    /// <see cref="Merge(string, string, MergeOptions?)"/>), fills it in place, and saves every part the
    /// merge is known to touch by name. Split out purely so the staging/atomic-replace plumbing above stays
    /// readable; this method owns the actual merge work and all of its saves.
    /// </summary>
    private static MergeResult MergeIntoOpenCopy(string stagingPath, MergeOptions options)
    {
        using var doc = WordprocessingDocument.Open(stagingPath, true);
        var schema = SchemaProvider.FromLayout(doc);

        // Backlog B23: pre-scan the (already-open) layout for every repeating section's own row binding
        // BEFORE generating sample data, so generation can multiply rows only for the data items the
        // document actually repeats - see ScanRepeaterConsumedPaths and SampleDataOptions.RepeaterConsumedPaths.
        var repeaterConsumedPaths = ScanRepeaterConsumedPaths(doc.MainDocumentPart!, schema);
        var data = SampleDataGenerator.Generate(schema, new SampleDataOptions
        {
            Seed = options.Seed,
            Rows = options.Rows,
            DataOverridesPath = options.DataOverridesPath,
            MaxRowsPerItem = options.MaxRowsPerRepeater,
            MaxTotalInstances = options.MaxTotalInstances,
            RepeaterConsumedPaths = repeaterConsumedPaths,
        });

        var result = Merge(doc, data, options);

        var main = doc.MainDocumentPart!;
        main.Document?.Save();
        foreach (var header in main.HeaderParts)
        {
            header.Header?.Save();
        }

        foreach (var footer in main.FooterParts)
        {
            footer.Footer?.Save();
        }

        // Only reached (part is loaded) when StripExternalRelationships actually touched the MAIN document's
        // settings.xml (e.g. to remove a dangling w:attachedTemplate element - the common real case, proven
        // end-to-end by McpHostToolTests' SalesInvoiceForSubscriptionBilling corpus test, whose attachedTemplate
        // lives exactly here). Explicit Save mirrors the Document/header/footer saves above for this one
        // named part; it is NOT the only thing making the strip's edits durable. This whole method runs
        // inside the `using var doc = WordprocessingDocument.Open(stagingPath, true)` above, whose default
        // AutoSave is true, so ANY OTHER part the stripper may have touched but this method has no named
        // reference to (a glossary document's own settings part, or anything reached only via the package
        // root - see ExternalRelationshipStripper.Strip(OpenXmlPackage)) is still persisted by
        // AutoSave-on-Dispose even without an explicit .Save() call here. That combination - explicit Save
        // where a named reference already exists, AutoSave as the backstop everywhere else - is deliberate,
        // not an oversight; it is NOT separately regression-tested for a non-main-document settings part
        // today (accepted gap, no corpus layout exercises one). Formerly this AutoSave
        // ran directly against the CALLER's outputPath, so an exception partway through this method could
        // flush a half-merged tree straight into it; it now runs against a throwaway STAGING copy instead
        // (see the caller, Merge(string, string, MergeOptions?)), so that same AutoSave behavior on a
        // mid-merge exception only ever touches a file that gets deleted, never outputPath itself.
        if (main.DocumentSettingsPart is { IsRootElementLoaded: true } settingsPart)
        {
            settingsPart.Settings?.Save();
        }

        result = AppendRequestedRowsExceedsCapWarningIfNeeded(result, options);
        result = AppendSampleDataCappedWarningIfNeeded(result, data, options);
        return AppendLabelsConventionHintIfNeeded(result, schema, options);
    }

    // ---- repeater-consumed-paths pre-scan ----

    /// <summary>
    /// Scans the main document plus every header/footer of <paramref name="main"/> for every
    /// <c>w15:repeatingSection</c> control, resolves each one's own row-binding xpath structurally against
    /// <paramref name="schema"/> (via <see cref="ResolveRepeaterTargetPath"/>), and returns the set of
    /// <see cref="DataItem.Path"/> values found — i.e. every data item the DOCUMENT actually repeats. Feeds
    /// <see cref="SampleDataOptions.RepeaterConsumedPaths"/> so <see cref="SampleDataGenerator"/> only
    /// multiplies rows for data items something in the document actually reads as a repeating section,
    /// instead of every repeating item in the schema regardless of whether any control ever looks at it.
    /// </summary>
    /// <remarks>
    /// Run BEFORE <see cref="SampleDataGenerator.Generate"/>, against the SAME already-open document
    /// <see cref="Merge(string, string, MergeOptions?)"/> otherwise only reads AFTER generation — a two-phase
    /// "scan the document, then generate" split, not a callback into generation itself: at the point
    /// generation used to run, the document was already open and parsed for its schema anyway (<see
    /// cref="SchemaProvider.FromLayout(WordprocessingDocument)"/>), so this scan is simply run first and its
    /// (tiny — one xpath string per repeater) result threaded into <see cref="SampleDataOptions"/> instead of
    /// generation needing any awareness of the document's own OOXML tree.
    /// <para>
    /// A repeater whose own xpath cannot be structurally resolved against <paramref name="schema"/> at all —
    /// a <c>//</c> descendant-axis binding (never seen in a real corpus layout; <see cref="XPathReanchor"/>
    /// already treats these as unsupported for re-anchoring too), or one naming a segment that is not a real
    /// child data item at that point in the schema — contributes nothing to the returned set. Either shape is
    /// already a malformed/unusual binding <see cref="LayoutValidator.Resolves"/> flags separately as an
    /// <c>xpath-resolves</c> validation error; leaving its target out of the consumed set costs nothing
    /// additional, since <see cref="MergeEngine.ProcessRepeater"/> would find no matching row for it in the
    /// generated data either way (the same reason nothing structurally resolves it here).
    /// </para>
    /// </remarks>
    internal static HashSet<string> ScanRepeaterConsumedPaths(MainDocumentPart main, DatasetTree schema)
    {
        var consumed = new HashSet<string>(StringComparer.Ordinal);

        foreach (var (root, _) in PartWalker.ContentParts(main))
        {
            foreach (var sdt in root.Descendants<SdtElement>())
            {
                var pr = sdt.GetFirstChild<SdtProperties>();
                if (pr is null || SdtInspector.Classify(pr) != SdtInspector.Classification.Repeater)
                {
                    continue;
                }

                var binding = SdtInspector.FindRepeaterBinding(pr);
                var rXPath = binding is null ? null : SdtInspector.Attr(binding, "xpath", OoxmlNames.W);
                if (string.IsNullOrWhiteSpace(rXPath))
                {
                    continue;
                }

                if (ResolveRepeaterTargetPath(rXPath, schema) is { } path)
                {
                    consumed.Add(path);
                }
            }
        }

        return consumed;
    }

    /// <summary>
    /// Structurally walks <paramref name="schema"/> from its root, one <see cref="BindingXPath.Segments"/>
    /// step at a time, following <see cref="DataItem.FindChildItem"/> — the same segment-by-segment approach
    /// <see cref="LayoutValidator.Resolves"/> uses for its <c>xpath-resolves</c> check, but requiring EVERY
    /// segment (including the last) to name a data ITEM, never a leaf column: a repeater's own row binding
    /// always targets a repeating collection, never a single field. Returns the resolved item's own
    /// <see cref="DataItem.Path"/>, or null when <paramref name="rXPath"/>'s first segment does not name
    /// <paramref name="schema"/>'s own root, or any later segment does not name a real child data item at
    /// that point in the tree (see <see cref="ScanRepeaterConsumedPaths"/>'s own remarks for why an
    /// unresolvable binding is simply left out rather than treated as an error here).
    /// </summary>
    private static string? ResolveRepeaterTargetPath(string rXPath, DatasetTree schema)
    {
        var segments = BindingXPath.Segments(rXPath);
        if (segments.Count == 0 || !string.Equals(segments[0], schema.Root.Name, StringComparison.Ordinal))
        {
            return null;
        }

        var node = schema.Root;
        for (var i = 1; i < segments.Count; i++)
        {
            var child = node.FindChildItem(segments[i]);
            if (child is null)
            {
                return null;
            }

            node = child;
        }

        return node.Path;
    }

    /// <summary>
    /// Appends one <c>sample-data-capped</c> <see cref="MergeWarning"/> when
    /// <see cref="SampleDataset.Truncated"/> is set — i.e. sample-data generation hit
    /// <see cref="MergeOptions.MaxTotalInstances"/> and stopped early, so the merged preview is deliberately
    /// PARTIAL (some repeating sections show fewer rows than <see cref="MergeOptions.Rows"/> requested, or none
    /// at all, deeper in the tree). A no-op (returns <paramref name="result"/> unchanged) when generation
    /// stayed within budget or when a real <see cref="MergeOptions.DataOverridesPath"/> dataset was used
    /// (never generated, so never truncated here). Preserves "no silent caps": a bounded preview always says so.
    /// </summary>
    private static MergeResult AppendSampleDataCappedWarningIfNeeded(MergeResult result, SampleDataset data, MergeOptions options)
    {
        if (!data.Truncated)
        {
            return result;
        }

        var warnings = new List<MergeWarning>(result.Warnings)
        {
            new MergeWarning
            {
                Kind = "sample-data-capped",
                Message = $"Sample-data generation reached the global cap of {options.MaxTotalInstances} "
                    + "business data-item instance(s) and stopped early, so this preview is PARTIAL: a deeply "
                    + $"nested layout multiplies instances across depth, so Rows={options.Rows} would have "
                    + "generated far more. Lower the 'rows' argument for a complete preview of a nested layout.",
            },
        };

        return new MergeResult { Stats = result.Stats, Warnings = warnings };
    }

    /// <summary>
    /// Appends one <c>labels-convention-hint</c> <see cref="MergeWarning"/> when the schema carries the
    /// well-known <c>&lt;Labels&gt;</c> data-item shape (a data item literally named "Labels" whose direct
    /// columns are the report's captions) but the ACTIVE <see cref="LabelConvention"/> does not classify
    /// them — i.e. <see cref="LabelConvention.LabelsDataItemName"/> is unset and at least one direct
    /// column lacks a recognized label suffix. Without the rule, those caption columns receive
    /// type-inferred FIELD samples, so a preview's column headings render as raw numbers/dates (the exact
    /// InventoryOrderDetails symptom from the 2026-07-31 preview sweep — garbage headers with nothing
    /// telling the caller the fix already exists). The convention deliberately stays OFF by default (see
    /// <see cref="LabelConvention"/>'s own remarks); this hint makes the knob discoverable exactly when it
    /// matters, instead of silently degrading the preview. Skipped when a real
    /// <see cref="MergeOptions.DataOverridesPath"/> dataset was supplied — its caption values are real, so
    /// nothing here is being mis-sampled.
    /// </summary>
    private static MergeResult AppendLabelsConventionHintIfNeeded(MergeResult result, DatasetTree schema, MergeOptions options)
    {
        if (!string.IsNullOrEmpty(options.DataOverridesPath) || LabelConvention.Current.LabelsDataItemName is not null)
        {
            return result;
        }

        var labelsItem = FindDataItem(schema.Root, "Labels");
        var unclassified = labelsItem?.Columns.Count(c => !LabelConvention.Current.IsLabelName(c.Name)) ?? 0;
        if (labelsItem is null || unclassified == 0)
        {
            return result;
        }

        var warnings = new List<MergeWarning>(result.Warnings)
        {
            new MergeWarning
            {
                Kind = "labels-convention-hint",
                Message = $"This layout's dataset has a '{labelsItem.Name}' data item with {unclassified} "
                    + "column(s) the active label convention does not classify as labels, so their sample "
                    + "values are type-inferred from the column NAMES (numbers/dates/codes) and any bound "
                    + "column headings in this preview may render as raw values instead of caption text. If "
                    + $"'{labelsItem.Name}' holds this report's captions (the common BC shape), set the "
                    + $"BCWL_LABELS_DATA_ITEM environment variable to '{labelsItem.Name}' on the MCP server "
                    + "and restart it to preview (and classify) them as labels.",
                Location = labelsItem.Path,
            },
        };

        return new MergeResult { Stats = result.Stats, Warnings = warnings };
    }

    /// <summary>Depth-first search for a data item named <paramref name="name"/> (ordinal).</summary>
    private static DataItem? FindDataItem(DataItem node, string name)
    {
        if (string.Equals(node.Name, name, StringComparison.Ordinal))
        {
            return node;
        }

        foreach (var child in node.Children)
        {
            if (FindDataItem(child, name) is { } found)
            {
                return found;
            }
        }

        return null;
    }

    /// <summary>
    /// Appends one document-level <c>row-cap</c> <see cref="MergeWarning"/> when <paramref name="options"/>.Rows
    /// itself exceeds <paramref name="options"/>.MaxRowsPerRepeater AND no <see cref="MergeOptions.DataOverridesPath"/>
    /// was used — see <see cref="Merge(string, string, MergeOptions?)"/>'s own remarks for why this is
    /// needed to preserve "no silent caps" once generation bounds itself to the same cap. A no-op (returns
    /// <paramref name="result"/> unchanged) otherwise.
    /// </summary>
    private static MergeResult AppendRequestedRowsExceedsCapWarningIfNeeded(MergeResult result, MergeOptions options)
    {
        if (!string.IsNullOrEmpty(options.DataOverridesPath) || options.Rows <= options.MaxRowsPerRepeater)
        {
            return result;
        }

        var warnings = new List<MergeWarning>(result.Warnings)
        {
            new MergeWarning
            {
                Kind = "row-cap",
                Message = $"The requested Rows ({options.Rows}) exceeds MaxRowsPerRepeater "
                    + $"({options.MaxRowsPerRepeater}); every repeating data item's generated sample data is "
                    + $"limited to {options.MaxRowsPerRepeater} instance(s) (see SampleDataOptions.MaxRowsPerItem).",
            },
        };

        return new MergeResult { Stats = result.Stats, Warnings = warnings };
    }

    /// <summary>
    /// Mutates an already-open document in place against a prepared dataset — no file I/O of its own;
    /// the caller owns opening and saving. Exposed for testability against a hand-built
    /// <see cref="SampleDataset"/>, and used internally by the file-based overload above.
    /// <paramref name="options"/> is only consulted for <see cref="MergeOptions.MaxRowsPerRepeater"/> here
    /// (every other option already shaped <paramref name="data"/> itself, via <see cref="SampleDataGenerator.Generate"/>,
    /// before this method ever runs); null (the default — every pre-existing caller of this overload never
    /// passed one) falls back to a plain <c>new MergeOptions()</c>, keeping the default cap of 100.
    /// </summary>
    internal static MergeResult Merge(WordprocessingDocument doc, SampleDataset data, MergeOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(doc);
        ArgumentNullException.ThrowIfNull(data);

        options ??= new MergeOptions();

        var main = doc.MainDocumentPart
            ?? throw new InvalidDataException("Layout has no main document part.");

        var nsmgr = BuildNamespaceManager(main, data.Namespace);
        var state = new MergeState { MaxRowsPerRepeater = options.MaxRowsPerRepeater };
        SeedDocumentWideIdCounters(main, state);

        foreach (var (root, hostPart, partName) in PartWalker.ContentPartsWithHost(main))
        {
            WalkElement(root, data.Xml, 0, null, hostPart, partName, nsmgr, data, state);
        }

        // The steps above fill each content control's VISIBLE run text and clone repeater rows in the OOXML.
        // But BC layout controls are LIVE data-bound: on open, Word re-syncs every control from the mapped
        // custom XML data part and re-evaluates each w15:repeatingSection against it. Since the merge never
        // populates that part (it writes run text, not data), Word would discard the merged text — showing the
        // field-name placeholder from the part — and collapse each repeater back to its single template row.
        // Severing the bindings (and the repeating-section markers, so the cloned rows survive as static rows)
        // turns the merged document into a self-contained snapshot of exactly what the merge produced, so a
        // rendered preview actually shows the sample values and every generated row. Opt-in (preview only):
        // the default logical-merge output keeps its bindings for validation/structural inspection.
        if (options.FlattenBindingsForRender)
        {
            FlattenLiveBindings(main);
        }

        // Security hardening (opt-in - see MergeOptions.StripExternalRelationships): removes anything a
        // renderer would dereference on open (attachedTemplate, linked images, linked OLE objects, mail-
        // merge/subDocument/movie references) before this merge result is handed to a converter - covers the
        // package root's own relationships too, not just the main document part's. Runs AFTER the walk/
        // flatten steps above (order is immaterial - relationships are independent of bound-field/repeater
        // content) so its warnings are simply appended alongside whatever the walk itself raised.
        if (options.StripExternalRelationships)
        {
            state.Warnings.AddRange(ExternalRelationshipStripper.Strip(doc));
        }

        return new MergeResult
        {
            Stats = new MergeStats
            {
                FieldsFilled = state.FieldsFilled,
                RepeatersExpanded = state.RepeatersExpanded,
                RowsGenerated = state.RowsGenerated,
                Unresolved = state.Unresolved,
                PicturesFilled = state.PicturesFilled,
            },
            Warnings = state.Warnings,
        };
    }

    /// <summary>
    /// Removes every live data binding (<c>w:dataBinding</c> / <c>w15:dataBinding</c>) and repeating-section
    /// marker (<c>w15:repeatingSection</c> / <c>w15:repeatingSectionItem</c>) from the document body and every
    /// header/footer, so the merged output no longer re-syncs from the (un-populated) custom XML part when a
    /// renderer opens it. The content controls themselves and their now-static content (merged run text, cloned
    /// rows, filled pictures) are left intact — only the "refresh me from the data part" links are cut.
    /// </summary>
    private static void FlattenLiveBindings(MainDocumentPart main)
    {
        foreach (var (root, _) in PartWalker.ContentParts(main))
        {
            FlattenPart(root);
        }
    }

    private static void FlattenPart(OpenXmlElement root)
    {
        foreach (var pr in root.Descendants<SdtProperties>().ToList())
        {
            var toRemove = pr.ChildElements
                .Where(e =>
                    (e.LocalName == "dataBinding" && (e.NamespaceUri == OoxmlNames.W || e.NamespaceUri == OoxmlNames.W15))
                    || ((e.LocalName == "repeatingSection" || e.LocalName == "repeatingSectionItem")
                        && e.NamespaceUri == OoxmlNames.W15))
                .ToList();

            foreach (var element in toRemove)
            {
                element.Remove();
            }
        }
    }

    // ---- tree walk ----

    private sealed class MergeState
    {
        public int FieldsFilled { get; set; }

        public int RepeatersExpanded { get; set; }

        public int RowsGenerated { get; set; }

        public int Unresolved { get; set; }

        public int PicturesFilled { get; set; }

        /// <summary>Snapshot of <see cref="MergeOptions.MaxRowsPerRepeater"/> for this merge (see <see cref="ProcessRepeater"/>).</summary>
        public int MaxRowsPerRepeater { get; set; } = 100;

        /// <summary>
        /// Next fresh id to assign to a cloned row's <c>wp:docPr</c> element (see
        /// <see cref="SeedDocumentWideIdCounters"/> and <see cref="RegenerateClonedIds"/>). Seeded ONCE per
        /// merge, above every pre-existing <c>wp:docPr/@id</c> in the WHOLE document (main + every
        /// header/footer), then incremented once per docPr this merge regenerates across EVERY row it
        /// clones, regardless of which repeater or which part — so ids stay unique document-wide even
        /// though OpenXmlValidator's own <c>Sem_UniqueAttributeValue</c> check for docPr is actually scoped
        /// PER PART (verified empirically: two DIFFERENT parts may legally repeat the same docPr id — each
        /// OOXML part is its own XML document, so a schema-level uniqueness constraint cannot see across
        /// parts). A single document-wide counter trivially satisfies the narrower per-part requirement too,
        /// with no extra bookkeeping to track "which part is this id already used in".
        /// </summary>
        public uint NextDocPrId { get; set; } = 1;

        /// <summary>Same idea as <see cref="NextDocPrId"/>, for <c>w:bookmarkStart</c>/<c>w:bookmarkEnd</c> id pairs.</summary>
        public long NextBookmarkId { get; set; }

        /// <summary>
        /// Every <c>w:bookmarkStart/@w:name</c> seen anywhere in the document before this merge's first row
        /// clone (seeded by <see cref="SeedDocumentWideIdCounters"/>), plus every name minted for a later
        /// clone since. Used only so a renamed bookmark (see <see cref="RegenerateClonedIds"/>) never
        /// collides with some unrelated bookmark elsewhere in the document — NOT required for
        /// OpenXmlValidator cleanliness (verified empirically: <c>Sem_UniqueAttributeValue</c> does not fire
        /// on a duplicate bookmark NAME, only a duplicate id); this is defense-in-depth against Word's own
        /// bookmark navigation/cross-reference resolution silently preferring the wrong same-named instance.
        /// </summary>
        public HashSet<string> BookmarkNames { get; } = new(StringComparer.Ordinal);

        /// <summary>
        /// The one placeholder <see cref="ImagePart"/> added so far per hosting part, keyed by reference
        /// identity (parts never override equality) so every picture control in the same part reuses the
        /// same relationship id instead of a new part being added per picture.
        /// </summary>
        public Dictionary<OpenXmlPart, ImagePart> PlaceholderImageParts { get; } = new();

        public List<MergeWarning> Warnings { get; } = new();
    }

    /// <summary>
    /// Recursively walks <paramref name="element"/>'s children, filling bound fields and expanding
    /// repeaters. <paramref name="context"/> is the current data context node — an <see cref="XDocument"/>
    /// at the top of each part, an <see cref="XElement"/> once inside a repeater row — and
    /// <paramref name="consumedSteps"/> is how many of a binding's leading raw XPath steps that context
    /// already represents (see <see cref="XPathReanchor"/>). <paramref name="contextXPath"/> is the
    /// enclosing repeater's own absolute binding XPath backing that same count (null at the top of a walk,
    /// where <paramref name="consumedSteps"/> is 0) — threaded down so <see cref="XPathReanchor.Remainder"/>
    /// can verify a candidate binding's dropped prefix actually matches this context structurally, not just
    /// in step count. <paramref name="hostPart"/> is the <see cref="OpenXmlPart"/> that owns this walk's
    /// part — <see cref="MainDocumentPart"/> for document.xml, else the specific
    /// <see cref="HeaderPart"/>/<see cref="FooterPart"/> — needed so a picture control found here gets its
    /// placeholder image added to the correct part. <paramref name="depth"/> (0 at the top of each part's
    /// walk) enforces <see cref="ResourceLimits.MaxElementNestingDepth"/>: unlike the schema tree (bounded
    /// once at construction, see <see cref="SchemaProvider.BuildNode"/>'s own remarks), a crafted
    /// document.xml/header/footer's own element nesting has no such upstream ceiling, so this hand-rolled
    /// recursive walk must bound itself — an uncatchable <see cref="StackOverflowException"/> would take the
    /// whole server down rather than fail one call.
    /// </summary>
    private static void WalkElement(
        OpenXmlElement element,
        XNode context,
        int consumedSteps,
        string? contextXPath,
        OpenXmlPart hostPart,
        string partName,
        XmlNamespaceManager nsmgr,
        SampleDataset data,
        MergeState state,
        int depth = 0)
    {
        if (depth > ResourceLimits.MaxElementNestingDepth)
        {
            throw ResourceLimits.DepthExceeded("Document element", ResourceLimits.MaxElementNestingDepth);
        }

        foreach (var child in element.ChildElements.ToList())
        {
            if (child is SdtElement sdt)
            {
                var pr = sdt.GetFirstChild<SdtProperties>();
                switch (SdtInspector.Classify(pr))
                {
                    case SdtInspector.Classification.Repeater:
                        ProcessRepeater(sdt, pr, context, consumedSteps, contextXPath, hostPart, partName, nsmgr, data, state, depth + 1);
                        break; // Rows already recursed into; do not also walk the un-expanded template.

                    case SdtInspector.Classification.RepeaterItem:
                        // Defensive only: a well-formed layout never reaches an item sdt here directly —
                        // ProcessRepeater's own descendant search consumes it. Stay transparent if it does.
                        WalkElement(sdt, context, consumedSteps, contextXPath, hostPart, partName, nsmgr, data, state, depth + 1);
                        break;

                    case SdtInspector.Classification.Picture:
                        ProcessPicture(sdt, hostPart, partName, state);
                        break; // A picture's content is just a drawing; nothing further to walk into.

                    case SdtInspector.Classification.Bound:
                        ProcessField(sdt, pr, context, consumedSteps, contextXPath, partName, nsmgr, data, state);
                        WalkElement(sdt, context, consumedSteps, contextXPath, hostPart, partName, nsmgr, data, state, depth + 1);
                        break;

                    default:
                        WalkElement(sdt, context, consumedSteps, contextXPath, hostPart, partName, nsmgr, data, state, depth + 1);
                        break;
                }
            }
            else
            {
                WalkElement(child, context, consumedSteps, contextXPath, hostPart, partName, nsmgr, data, state, depth + 1);
            }
        }
    }

    // ---- repeater expansion ----

    private static void ProcessRepeater(
        SdtElement repeaterSdt,
        SdtProperties? pr,
        XNode context,
        int consumedSteps,
        string? contextXPath,
        OpenXmlPart hostPart,
        string partName,
        XmlNamespaceManager nsmgr,
        SampleDataset data,
        MergeState state,
        int depth)
    {
        if (pr is null)
        {
            return;
        }

        var binding = SdtInspector.FindRepeaterBinding(pr);
        var rXPath = binding is null ? null : SdtInspector.Attr(binding, "xpath", OoxmlNames.W);

        if (string.IsNullOrWhiteSpace(rXPath))
        {
            return;
        }

        // The row template is the sole repeatingSectionItem sdt among this repeater's descendants whose
        // NEAREST enclosing repeater is itself. Searching descendants (not just direct children) tolerates
        // an intervening wrapper sdt — real BC layouts sometimes wrap the item in a locked w:group — while
        // the nearest-ancestor filter still excludes a nested distinct repeater's own item, mirroring
        // LayoutValidator's repeater-shape invariant.
        var template = repeaterSdt.Descendants<SdtElement>()
            .FirstOrDefault(s => SdtInspector.IsRepeaterItem(s) && SdtInspector.NearestRepeaterAncestor(s) == repeaterSdt);

        if (template is null)
        {
            return;
        }

        List<XElement> rows;
        var remainder = XPathReanchor.Remainder(rXPath, consumedSteps, contextXPath);
        if (remainder is null)
        {
            // Either no re-anchoring is possible at all (top-level, `//` axis, bad step count), or —
            // shape (c) — this repeater is nested inside an enclosing repeater's row
            // but its own binding is structurally divergent from that row (equal step count, different
            // element(s)): dropping the prefix would silently pick the wrong context node and, worse, hand
            // "." to SafeSelectElements below, cloning one bogus row from the row node itself. Evaluating
            // the full absolute XPath from the document root instead always resolves the REAL matched
            // node(s), however many that turns out to be.
            rows = SafeSelectElements(data.Xml, rXPath, nsmgr, partName, rXPath, state);
            state.Warnings.Add(new MergeWarning
            {
                Kind = "xpath-fallback",
                Message = "Repeater XPath could not be re-anchored to its row context; evaluated from the "
                    + $"document root instead: '{rXPath}'.",
                Location = $"{partName}: {rXPath}",
            });
        }
        else
        {
            var expression = remainder.Length == 0 ? "." : remainder;
            rows = SafeSelectElements(context, expression, nsmgr, partName, rXPath, state);
        }

        // Robustness cap: a SINGLE repeating section (this call — a nested repeater's own
        // expansion, once per surviving clone of its enclosing row, is a separate call and gets the SAME
        // cap applied independently, not a shared global budget) never clones more than MaxRowsPerRepeater
        // rows, regardless of whether the excess came from a large Rows value or from an unexpectedly large
        // number of matching rows in a real DataOverridesPath dataset. Never silent: a row-cap warning
        // always names the repeater and the cap whenever this actually trims anything.
        if (rows.Count > state.MaxRowsPerRepeater)
        {
            var matchedCount = rows.Count;
            rows = rows.Take(state.MaxRowsPerRepeater).ToList();
            state.Warnings.Add(new MergeWarning
            {
                Kind = "row-cap",
                Message = $"Repeater '{rXPath}' matched {matchedCount} row(s), exceeding the "
                    + $"MaxRowsPerRepeater cap of {state.MaxRowsPerRepeater}; only the first "
                    + $"{state.MaxRowsPerRepeater} were generated.",
                Location = $"{partName}: {rXPath}",
            });
        }

        // Use the SAME raw-step splitter as XPathReanchor.Remainder (not BindingXPath.Segments, which
        // strips prefixes/predicates for a different purpose) so the count of steps this row's children
        // should treat as "already consumed" can never drift from what Remainder itself dropped. This
        // repeater's OWN xpath becomes the new contextXPath for everything inside its rows — regardless of
        // whether ITS OWN remainder above had to fall back to root evaluation, rXPath still absolutely and
        // correctly describes rowNode's location, so descendants re-anchoring against it is exactly right.
        var childConsumedSteps = XPathReanchor.RawSteps(rXPath).Count;

        for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            var rowNode = rows[rowIndex];

            // Cloned rows intentionally keep the template's original SDT w:id values un-renumbered: w:id
            // has no schema-level uniqueness constraint, and OpenXmlValidator passes clean on the full
            // corpus with SDT id duplicates present. wp:docPr and bookmark w:id ARE schema-uniqueness-
            // constrained though (Sem_UniqueAttributeValue), so RegenerateClonedIds below
            // gives every clone its own fresh values for exactly those two families.
            var clone = template.CloneNode(true);
            template.InsertBeforeSelf(clone);
            RegenerateClonedIds(clone, state, rowIndex: rowIndex + 1);
            WalkElement(clone, rowNode, childConsumedSteps, rXPath, hostPart, partName, nsmgr, data, state, depth);
        }

        template.Remove();

        state.RepeatersExpanded++;
        state.RowsGenerated += rows.Count;
    }

    // ---- cloned-row id regeneration ----

    /// <summary>
    /// Scans every content part (main document + every header/footer — the same set <see cref="ProcessRepeater"/>
    /// ever clones a row into; see <see cref="PartWalker.ContentParts"/>) ONCE, before this merge clones its
    /// first row, for the highest pre-existing <c>wp:docPr/@id</c> and <c>w:bookmarkStart/@w:id</c> in the
    /// WHOLE document, plus every existing bookmark name — so <see cref="RegenerateClonedIds"/> can hand
    /// out ids/names that never collide with anything already there. Read-only; mutates only
    /// <paramref name="state"/>. Non-numeric bookmark ids (never seen in practice — BC/Word always emit a
    /// plain decimal — but nothing in the schema forbids one) are simply skipped for the max computation
    /// rather than throwing, since a merge must never fail outright over cosmetic id bookkeeping.
    /// </summary>
    private static void SeedDocumentWideIdCounters(MainDocumentPart main, MergeState state)
    {
        uint maxDocPrId = 0;
        long maxBookmarkId = -1;

        foreach (var (root, _) in PartWalker.ContentParts(main))
        {
            foreach (var docPr in root.Descendants<DrawWordprocessing.DocProperties>())
            {
                if (docPr.Id?.Value is { } id && id > maxDocPrId)
                {
                    maxDocPrId = id;
                }
            }

            foreach (var bookmarkStart in root.Descendants<BookmarkStart>())
            {
                if (!string.IsNullOrEmpty(bookmarkStart.Name?.Value))
                {
                    state.BookmarkNames.Add(bookmarkStart.Name!.Value!);
                }

                if (long.TryParse(
                        bookmarkStart.Id?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var bId)
                    && bId > maxBookmarkId)
                {
                    maxBookmarkId = bId;
                }
            }
        }

        // Saturating rather than `maxDocPrId + 1`: a document already carrying a wp:docPr id of
        // uint.MaxValue (astronomically unrealistic for a real BC layout — it would need over 4 billion
        // distinct docPr elements) would otherwise silently wrap the seed itself to 0, colliding with a
        // real id. Saturating here only defers the theoretical problem to the very first increment inside
        // RegenerateClonedIds (which is NOT similarly guarded) rather than eliminating it outright — judged
        // acceptable given how far outside any real corpus/BC-generated layout this scenario sits.
        state.NextDocPrId = maxDocPrId == uint.MaxValue ? uint.MaxValue : maxDocPrId + 1;
        state.NextBookmarkId = maxBookmarkId + 1;
    }

    /// <summary>
    /// Walks a just-cloned repeater row (see <see cref="ProcessRepeater"/>) and gives every <c>wp:docPr</c>
    /// element (DrawingML picture/graphic-frame identity) a fresh, merge-wide-unique id from
    /// <see cref="MergeState.NextDocPrId"/>, and every <c>w:bookmarkStart</c>/<c>w:bookmarkEnd</c> a fresh,
    /// merge-wide-unique id from <see cref="MergeState.NextBookmarkId"/>. A start/end PAIR is matched by
    /// their SHARED OLD id (captured before either end is touched) and given the SAME new id, so the pair
    /// stays a pair after renumbering — but ONLY when BOTH ends of the bookmark live inside THIS clone. A
    /// bookmark whose range SPANS the repeater boundary (one end inside the cloned row, the other outside
    /// it entirely — before or after the repeater) cannot stay paired: the outside end exists exactly once
    /// regardless of <c>Rows</c>, while the inside end is duplicated once per row, so there is no single new
    /// id that could keep both sides matched across every clone. Rather than walking the WHOLE document to
    /// reconnect such a range (considered and rejected as disproportionate for a preview/validation output —
    /// no corpus layout has ever been observed to place a bookmark that way), the inside end is simply given
    /// a FRESH id of its own, deliberately ORPHANING it (its id no longer matches anything). This is exactly
    /// what already happens, unconditionally, to a <see cref="BookmarkStart"/> whose matching
    /// <see cref="BookmarkEnd"/> lives outside the clone — confirmed validator-clean either way, since
    /// <c>Sem_UniqueAttributeValue</c> only requires ids to be UNIQUE, not that every end has a matching
    /// start somewhere. Both attributes are exactly what OpenXmlValidator's <c>Sem_UniqueAttributeValue</c>
    /// flags once a row containing either is cloned more than once; other candidate id
    /// families were checked and found NOT to need this (see remarks below).
    /// </summary>
    /// <param name="rowIndex">
    /// This row's 1-based position among the rows THIS repeater expansion is generating (not a document-
    /// wide row count). The first row keeps every bookmark's ORIGINAL name unchanged — preserving any
    /// external reference to it, and matching the pre-fix behavior for the common single-row case — while
    /// every subsequent row's bookmark names get a deterministic <c>"_r{rowIndex}"</c> suffix, de-duplicated
    /// against <see cref="MergeState.BookmarkNames"/> in the rare case that also collides. This is purely
    /// defensive: verified empirically that OpenXmlValidator does NOT flag a duplicate bookmark NAME, only
    /// a duplicate id, so a merge would already be validator-clean without it.
    /// </param>
    /// <remarks>
    /// Other candidate <c>Sem_UniqueAttributeValue</c> sources were checked against what real BC layouts
    /// (the tests/corpus fixtures) and the OOXML schema actually contain and found unnecessary here:
    /// <c>wp14:anchorId</c>/<c>wp14:editId</c> (verified empirically — duplicates raise no validator error
    /// at all); <c>w:permStart</c>/<c>w:permEnd</c> range permissions and tracked-change markers
    /// (<c>w:ins</c>/<c>w:del</c>/<c>w:moveFrom*</c>/<c>w:moveTo*</c>) — absent from every corpus layout, and
    /// not something a BC-generated report template would ever carry. <c>w:commentRangeStart</c>/
    /// <c>w:commentRangeEnd</c>/<c>w:commentReference</c> ids are ALSO absent from every corpus layout, and
    /// were left unhandled for a structural reason, not just rarity: a comment's id also has to match a
    /// <c>w:comment/@w:id</c> in the separate <c>comments.xml</c> part, which this merge never touches — a
    /// clone containing a comment would need a whole new comment part entry, not just a renumbered
    /// reference, which is out of scope for a fix aimed at the two id families actually exercised by the
    /// corpus.
    /// </remarks>
    private static void RegenerateClonedIds(OpenXmlElement clone, MergeState state, int rowIndex)
    {
        foreach (var docPr in clone.Descendants<DrawWordprocessing.DocProperties>().ToList())
        {
            docPr.Id = state.NextDocPrId++;
        }

        // Collected up front (not re-queried per bookmarkStart) so the "still unmatched after the loop
        // below" check works: a BookmarkEnd whose matching BookmarkStart lives OUTSIDE this clone (the
        // range spans the repeater boundary — see this method's own remarks) is never reached by that loop
        // at all, so without the second pass below it would keep the template's original id verbatim in
        // EVERY clone — N identical ids, the exact Sem_UniqueAttributeValue error this method exists to
        // prevent. Tracked by parallel index (not a HashSet/Remove keyed on the element itself) so this
        // never depends on OpenXmlElement's equality semantics.
        var endsInClone = clone.Descendants<BookmarkEnd>().ToList();
        var endMatched = new bool[endsInClone.Count];

        foreach (var bookmarkStart in clone.Descendants<BookmarkStart>().ToList())
        {
            var oldId = bookmarkStart.Id?.Value;
            var newId = (state.NextBookmarkId++).ToString(CultureInfo.InvariantCulture);

            if (oldId is not null)
            {
                for (var i = 0; i < endsInClone.Count; i++)
                {
                    if (!endMatched[i] && endsInClone[i].Id?.Value == oldId)
                    {
                        endsInClone[i].Id = newId;
                        endMatched[i] = true;
                        break;
                    }
                }
            }

            bookmarkStart.Id = newId;

            var originalName = bookmarkStart.Name?.Value;
            if (rowIndex > 1 && !string.IsNullOrEmpty(originalName))
            {
                var candidate = $"{originalName}_r{rowIndex}";
                var suffix = 0;
                while (!state.BookmarkNames.Add(candidate))
                {
                    suffix++;
                    candidate = $"{originalName}_r{rowIndex}_{suffix}";
                }

                bookmarkStart.Name = candidate;
            }
        }

        // Any BookmarkEnd left unmatched above has its matching BookmarkStart outside this clone - give it
        // a fresh id of its own (deliberately orphaning the range) rather than leaving N clones sharing the
        // template's original id.
        for (var i = 0; i < endsInClone.Count; i++)
        {
            if (!endMatched[i])
            {
                endsInClone[i].Id = (state.NextBookmarkId++).ToString(CultureInfo.InvariantCulture);
            }
        }
    }

    // ---- field fill ----

    private static void ProcessField(
        SdtElement fieldSdt,
        SdtProperties? pr,
        XNode context,
        int consumedSteps,
        string? contextXPath,
        string partName,
        XmlNamespaceManager nsmgr,
        SampleDataset data,
        MergeState state)
    {
        if (pr is null)
        {
            return;
        }

        var binding = SdtInspector.FindBinding(pr);
        var xpath = binding is null ? null : SdtInspector.Attr(binding, "xpath", OoxmlNames.W);

        if (string.IsNullOrWhiteSpace(xpath))
        {
            return;
        }

        XElement? resolved;
        var remainder = XPathReanchor.Remainder(xpath, consumedSteps, contextXPath);
        if (remainder is null)
        {
            // Either no re-anchoring is possible at all, or — shapes (a)/(b) — this
            // field sits inside a repeater row but its own binding is structurally divergent from that
            // row's own path: an equal-depth sibling (e.g. a Header-level field bound inside a same-depth
            // Line row) would otherwise have its prefix dropped as if it matched, resolving to the ROW node
            // and rendering its concatenated text; a deeper-but-divergent path would evaluate its tail
            // against the wrong node entirely. Falling back to the full absolute XPath from the document
            // root always resolves the value BC itself would show.
            resolved = SafeSelectElement(data.Xml, xpath, nsmgr, partName, xpath, state);
            state.Warnings.Add(new MergeWarning
            {
                Kind = "xpath-fallback",
                Message = "Field XPath could not be re-anchored to its row context; evaluated from the "
                    + $"document root instead: '{xpath}'.",
                Location = $"{partName}: {xpath}",
            });
        }
        else if (remainder.Length == 0)
        {
            // The binding targets the context node itself (its own XPath is fully consumed already).
            resolved = context as XElement ?? (context as XDocument)?.Root;
        }
        else
        {
            resolved = SafeSelectElement(context, remainder, nsmgr, partName, xpath, state);
        }

        if (resolved is null)
        {
            var leaf = BindingXPath.LeafName(xpath) ?? "field";
            SetSdtText(fieldSdt, $"«{leaf}?»");
            state.Warnings.Add(new MergeWarning
            {
                Kind = "unresolved-binding",
                Message = $"Binding XPath did not resolve against the sample dataset: '{xpath}'.",
                Location = $"{partName}: {xpath}",
            });
            state.Unresolved++;
            return;
        }

        if (SetSdtText(fieldSdt, resolved.Value))
        {
            state.FieldsFilled++;
        }
        else
        {
            state.Warnings.Add(new MergeWarning
            {
                Kind = "content-write-failed",
                Message = $"Resolved a value for binding '{xpath}' but the control has no writable text content.",
                Location = $"{partName}: {xpath}",
            });
        }
    }

    /// <summary>
    /// Sets an sdt's visible text by writing the first <c>w:t</c> descendant of its content and blanking
    /// the rest, preserving paragraph/run/rPr formatting instead of replacing content wholesale. If the
    /// content has no <c>w:t</c> at all, a run is appended to the first paragraph found in it. Clears
    /// <c>w:showingPlcHdr</c> so Word stops treating the content as placeholder text. Returns false (and
    /// writes nothing) when the sdt has no content element, or has neither a <c>w:t</c> nor a paragraph to
    /// host a new run — callers must not count that as a fill.
    /// </summary>
    private static bool SetSdtText(SdtElement sdt, string value)
    {
        var pr = sdt.GetFirstChild<SdtProperties>();
        if (pr is not null)
        {
            SdtInspector.FirstChild(pr, "showingPlcHdr", OoxmlNames.W)?.Remove();
        }

        var content = SdtInspector.FirstChild(sdt, "sdtContent", OoxmlNames.W);
        if (content is null)
        {
            return false;
        }

        var texts = content.Descendants<Text>().ToList();
        if (texts.Count == 0)
        {
            var paragraph = content.Descendants<Paragraph>().FirstOrDefault();
            if (paragraph is null)
            {
                return false;
            }

            paragraph.AppendChild(new Run(new Text(value) { Space = SpaceProcessingModeValues.Preserve }));
            return true;
        }

        texts[0].Text = value;
        texts[0].Space = SpaceProcessingModeValues.Preserve;
        for (var i = 1; i < texts.Count; i++)
        {
            texts[i].Text = string.Empty;
        }

        return true;
    }

    // ---- picture fill ----

    /// <summary>
    /// Repoints a picture control's blip to the shared placeholder <see cref="ImagePart"/> for
    /// <paramref name="hostPart"/> (see <see cref="GetOrCreatePlaceholderImagePart"/>), so a downstream
    /// converter rasterizes a visible gray box instead of BC's real image reference — in an
    /// unmerged layout, a 10-byte stub. Looks for a blip anywhere among the sdt's descendants, tolerating
    /// whatever <c>a:blipFill</c>/<c>pic:pic</c> nesting the drawing uses, and leaves the control
    /// untouched (raising a <c>picture-no-blip</c> warning, never throwing) when there is no blip or no
    /// <c>r:embed</c> relationship to repoint. Any other relationship — including the original stub image
    /// part — is left exactly as-is; only the blip's own <c>r:embed</c> attribute changes.
    /// </summary>
    private static void ProcessPicture(SdtElement pictureSdt, OpenXmlPart hostPart, string partName, MergeState state)
    {
        var blip = pictureSdt.Descendants<Blip>().FirstOrDefault();
        if (blip is null || string.IsNullOrEmpty(blip.Embed?.Value))
        {
            state.Warnings.Add(new MergeWarning
            {
                Kind = "picture-no-blip",
                Message = "Picture control has no blip with an embed relationship to repoint; left untouched.",
                Location = partName,
            });
            return;
        }

        var imagePart = GetOrCreatePlaceholderImagePart(hostPart, state);
        blip.Embed = hostPart.GetIdOfPart(imagePart);
        state.PicturesFilled++;
    }

    /// <summary>
    /// Returns the placeholder <see cref="ImagePart"/> for <paramref name="hostPart"/>, adding (and
    /// caching in <see cref="MergeState.PlaceholderImageParts"/>) exactly one the first time a picture in
    /// that part needs it, so every picture sharing the same hosting part reuses the same relationship id
    /// rather than growing one image part per control. Uses the generic
    /// <see cref="OpenXmlPartContainer.AddNewPart{T}(string, string?)"/> — not the concrete
    /// <c>AddImagePart</c> convenience overloads, which are code-generated per container-part type
    /// (<see cref="MainDocumentPart"/>, <see cref="HeaderPart"/>, <see cref="FooterPart"/> each have their
    /// own) rather than declared on the common <see cref="OpenXmlPart"/> base this method is handed —
    /// so this works no matter which kind of part is currently being walked.
    /// </summary>
    private static ImagePart GetOrCreatePlaceholderImagePart(OpenXmlPart hostPart, MergeState state)
    {
        if (state.PlaceholderImageParts.TryGetValue(hostPart, out var existing))
        {
            return existing;
        }

        var imagePart = hostPart.AddNewPart<ImagePart>("image/png");
        using (var stream = new MemoryStream(PlaceholderImage.PngBytes))
        {
            imagePart.FeedData(stream);
        }

        state.PlaceholderImageParts.Add(hostPart, imagePart);
        return imagePart;
    }

    // ---- xpath evaluation (never throws across the merge boundary) ----

    /// <summary>
    /// Evaluates <paramref name="expression"/> against <paramref name="node"/> and returns every matching
    /// element, or an empty list plus an <c>xpath-error</c> warning if the expression is malformed enough
    /// to raise <see cref="XPathException"/> — a merge must never fail outright over one bad binding.
    /// </summary>
    private static List<XElement> SafeSelectElements(
        XNode node, string expression, XmlNamespaceManager nsmgr, string partName, string originalXPath, MergeState state)
    {
        try
        {
            return node.XPathSelectElements(expression, nsmgr).ToList();
        }
        catch (XPathException ex)
        {
            state.Warnings.Add(new MergeWarning
            {
                Kind = "xpath-error",
                Message = $"XPath evaluation failed for '{originalXPath}': {ex.Message}",
                Location = $"{partName}: {originalXPath}",
            });
            return new List<XElement>();
        }
    }

    /// <summary>Single-result counterpart of <see cref="SafeSelectElements"/>; see its remarks.</summary>
    private static XElement? SafeSelectElement(
        XNode node, string expression, XmlNamespaceManager nsmgr, string partName, string originalXPath, MergeState state)
    {
        try
        {
            return node.XPathSelectElement(expression, nsmgr);
        }
        catch (XPathException ex)
        {
            state.Warnings.Add(new MergeWarning
            {
                Kind = "xpath-error",
                Message = $"XPath evaluation failed for '{originalXPath}': {ex.Message}",
                Location = $"{partName}: {originalXPath}",
            });
            return null;
        }
    }

    // ---- namespace manager ----

    /// <summary>
    /// Maps prefix <c>ns0</c> to the dataset namespace, then scans every <c>w:dataBinding</c> /
    /// <c>w15:dataBinding</c>'s <c>w:prefixMappings</c> across document.xml + headers + footers and
    /// registers any additional prefixes found, so XPath evaluation is robust to layouts that used a
    /// prefix other than <c>ns0</c>.
    /// </summary>
    private static XmlNamespaceManager BuildNamespaceManager(MainDocumentPart main, string dataNamespace)
    {
        var nsmgr = new XmlNamespaceManager(new NameTable());
        nsmgr.AddNamespace("ns0", dataNamespace);
        var seen = new HashSet<string>(StringComparer.Ordinal) { "ns0" };

        foreach (var (root, _) in PartWalker.ContentParts(main))
        {
            foreach (var pr in root.Descendants<SdtProperties>())
            {
                var binding = SdtInspector.FindBinding(pr);
                if (binding is null)
                {
                    continue;
                }

                var mappings = SdtInspector.Attr(binding, "prefixMappings", OoxmlNames.W);
                if (string.IsNullOrEmpty(mappings))
                {
                    continue;
                }

                foreach (Match m in PrefixMappingPattern.Matches(mappings))
                {
                    var prefix = m.Groups["prefix"].Value;
                    if (seen.Add(prefix))
                    {
                        nsmgr.AddNamespace(prefix, m.Groups["uri"].Value);
                    }
                }
            }
        }

        return nsmgr;
    }
}
