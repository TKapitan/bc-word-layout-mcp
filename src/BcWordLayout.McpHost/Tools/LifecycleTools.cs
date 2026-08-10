using System.ComponentModel;
using System.Text.Json;
using BcWordLayout.Domain;
using BcWordLayout.Merge;
using BcWordLayout.Render;
using ModelContextProtocol;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using static BcWordLayout.McpHost.Tools.ToolGuards;

namespace BcWordLayout.McpHost.Tools;

/// <summary>
/// MCP tools spanning a layout's whole lifecycle rather than a single in-place content edit: creating a
/// brand-new layout from a schema (<c>create_layout</c>), swapping an existing layout's BC dataset part for
/// a new schema (<c>refresh_xml_part</c>), and rendering a mock PDF preview (<c>preview_layout</c>).
/// <c>refresh_xml_part</c> is still a MUTATING tool and routes through <see cref="ToolGuards.GuardRefresh"/>
/// (a thin <see cref="ToolGuards.GuardMutate{TResult}"/> instantiation, same save-or-reject safety as every
/// other mutating tool); <c>preview_layout</c> never mutates the ORIGINAL layout (it merges into a separate
/// working copy) but still takes the SAME per-path lock PAIR every mutating tool does — the in-process
/// <see cref="ToolGuards.EditLockFor"/> lock THEN, inside it, the cross-process
/// <see cref="CrossProcessLock"/> mutex — so a preview can never
/// observe a half-written file mid-edit, nor race a concurrent preview/edit of the same layout FROM ANOTHER
/// HOST PROCESS either; <c>create_layout</c> writes a brand-new file, so it takes only the lock PAIR keyed
/// on its <c>outputPath</c> (see its own remarks) rather than the full <see cref="ToolGuards.GuardMutate{TResult}"/>
/// choreography, which exists for editing an EXISTING file in place.
/// </summary>
[McpServerToolType]
public static class LifecycleTools
{
    /// <summary>
    /// Resolves the <see cref="IPdfConverter"/> <c>preview_layout</c> uses for a given
    /// <see cref="PdfConverterKind"/>. Defaults to the real <see cref="PdfConverterFactory.Select"/>, so
    /// production behavior is unchanged and this seam is invisible outside tests. A test substitutes a fake
    /// converter here (see <c>FakePdfConverter</c> in <c>BcWordLayout.Tests</c>) so <c>preview_layout</c>'s
    /// PDF-conversion outcome handling (<c>converterAvailable</c>/<c>conversionOk</c>/<c>conversionError</c>/
    /// <c>pdfPath</c> propagation) can be asserted deterministically without depending on whether Word or
    /// LibreOffice is actually installed on the machine running the suite (DI was
    /// deliberately never introduced elsewhere - this is the one, minimal, static seam that needed one).
    /// Internal rather than private specifically so a test can swap it, via the same <c>InternalsVisibleTo</c>
    /// grant in <c>BcWordLayout.McpHost.csproj</c> that already exposes <see cref="PreviewPathHash"/> etc.
    /// <para>
    /// TEST ISOLATION: this is a single process-wide static, so every test that swaps it MUST restore it to
    /// <see cref="PdfConverterFactory.Select"/> in a <c>finally</c> block before returning, AND every test
    /// CLASS with a test that calls <c>LifecycleTools.PreviewLayout</c> at all MUST carry
    /// <c>[Collection("preview-converter-seam")]</c> (see <c>PreviewConverterSeamCollection</c> in
    /// <c>BcWordLayout.Tests</c>, including the precise reach analysis behind the rule). xUnit runs classes
    /// in the same named collection sequentially but distinct
    /// collections in parallel, so a seam-swapping test in one class cannot otherwise be fenced from a
    /// seam-READING test in another - a leaked fake mid-swap would fail it intermittently. The collection
    /// membership, not class co-location, is the invariant that keeps swapping this static race-free.
    /// </para>
    /// </summary>
    internal static Func<PdfConverterKind, IPdfConverter> SelectConverter { get; set; } = PdfConverterFactory.Select;

    [McpServerTool(Name = "create_layout")]
    [Description("Create a NEW blank BC Word report layout .docx from a report schema, ready to be edited "
                 + "with insert_field/insert_label/insert_repeater_table. schemaSource is either an existing "
                 + ".docx layout (its BC dataset custom XML part is copied verbatim) or a standalone exported "
                 + "schema .xml. Optionally starts from templatePath instead of a blank document - this MUST "
                 + "be an UNBOUND branded/styled .docx shell (headers/footers, logo, fonts/styles), NOT a "
                 + "full BC layout with its own bound content controls: the template's own body is kept as-is "
                 + "(a heading naming the report is appended only when the body is entirely empty) and its "
                 + "pre-existing bound controls, if any, are never stripped or rebound. If the template "
                 + "already has its own BC dataset part it is removed and replaced (reported as "
                 + "replacedExistingBcPart=true) - if that leaves any of the template's own bound controls "
                 + "stale against the fresh storeItemID, the call FAILS outright with error code "
                 + "template_not_unbound rather than silently shipping a broken layout (a template with a BC "
                 + "part but zero bound controls of its own is unaffected and still succeeds). The created "
                 + "layout always ships exactly one BC custom XML part with a freshly generated storeItemID "
                 + "plus the glossary part the insert_* tools' placeholders depend on; a BLANK (non-template) "
                 + "layout also ships an empty header and footer part wired into its page setup, so "
                 + "insert_field/insert_label can target layoutPart='header'/'footer' straight away, AND a "
                 + "default styles part pinning its typography (Calibri 11pt docDefaults plus the standard "
                 + "Normal/TableGrid style definitions) so a from-scratch layout renders with the same font "
                 + "in Word and in Business Central instead of each renderer's own application default, AND a "
                 + "settings part declaring compatibilityMode 15 - the mode every stock BC layout declares "
                 + "and the one in which repeating-section controls exist, so opening the layout in Word and "
                 + "saving it does not convert its repeaters to plain rich-text controls and drop their "
                 + "bindings (a template keeps its own headers/footers, styles/theme AND settings untouched - "
                 + "templatePath is the way to pin custom branding; validate_layout's compatibility-mode "
                 + "check warns if a template shell leaves the layout below mode 15). Its own post-build "
                 + "quick validation travels with the result - check quickValidation.passed rather than "
                 + "assuming Ok=true means every possible issue was surfaced as data (a successful call is "
                 + "guaranteed free of the stale-binding damage above, but a template's pre-existing content "
                 + "can still carry its own warning-level finding). outputPath is overwritten if it already "
                 + "exists; its parent directory is created if missing; the write is atomic (built in a temp "
                 + "file next to outputPath first, so a failure - including the template_not_unbound refusal "
                 + "- never leaves outputPath partially written or written at all). "
                 + "CONCURRENCY: outputPath is locked (the same in-process + cross-process lock pair every "
                 + "editing tool uses) for the duration of the build, so two calls that happen to target the "
                 + "SAME outputPath - even from two different host processes - can never interleave their "
                 + "writes; a lock that cannot be acquired within the timeout fails with file_locked rather "
                 + "than corrupting or silently losing one of the two builds.")]
    public static ToolResponse CreateLayout(
        [Description("Absolute path to an existing .docx layout (its BC dataset part is reused) OR a "
                     + "standalone schema .xml file.")] string schemaSource,
        [Description("Absolute path to write the new layout .docx to.")] string outputPath,
        [Description("Optional absolute path to an UNBOUND branded/styled .docx to start from instead of a "
                     + "blank document (e.g. a letterhead template with headers/footers/logo/styles but no "
                     + "bound content controls of its own - see this tool's own description for why).")]
        string? templatePath = null,
        [Description("Optional text for the heading paragraph a BLANK (non-template) layout starts with. "
                     + "Omit for the default (the report's own name, e.g. 'Standard_Sales_Order_Conf'); "
                     + "pass a human title (e.g. 'Sales Order Confirmation') when authoring a document from "
                     + "scratch; pass an empty string for NO heading at all (the body starts with one empty "
                     + "paragraph). Ignored when the template's body already has content.")]
        string? headingText = null)
    {
        return Guard(() =>
        {
            // A null/blank outputPath has no real path to lock on yet - EditLockFor/CrossProcessLock.TryAcquire
            // both call Path.GetFullPath, which throws its OWN differently-worded ArgumentException (with a
            // generic/unhelpful ParamName) for a blank string, pre-empting LayoutBuilder.Create's own "Output
            // path must not be empty." check below and silently degrading InvalidArgumentHint's "outputpath"
            // branch to the generic fallback. Thrown HERE, with the exact same message/ParamName
            // LayoutBuilder.Create itself uses, so that never happens (see InvalidArgumentHintCoverageTests'
            // CreateLayout_empty_outputPath test, which pinned this before any locking existed here).
            if (string.IsNullOrWhiteSpace(outputPath))
            {
                throw new ArgumentException("Output path must not be empty.", nameof(outputPath));
            }

            // outputPath itself is a fresh file, never opened by any OTHER tool while it doesn't yet exist -
            // so this needs neither GuardMutate's copy/validate/atomic-swap choreography (there is nothing
            // existing to protect a working copy OF) nor GuardRead's coordination with a rename-in-progress.
            // It still needs the lock PAIR ("outputPath could collide with
            // another process's edit"): two create_layout calls racing the SAME outputPath - plausible if an
            // agent retries a call it believes failed, or two IDE windows both scaffold the same new file -
            // would otherwise interleave LayoutBuilder.Create's own temp-file-build-then-rename sequences
            // against each other. Locked for the whole build, same as GuardMutate locks for the whole edit.
            lock (EditLockFor(outputPath))
            {
                using var crossLock = CrossProcessLock.TryAcquire(outputPath, CrossProcessLock.MutatingTimeout);
                if (!crossLock.Acquired)
                {
                    return CrossProcessLockTimeoutFailure(outputPath, CrossProcessLock.MutatingTimeout);
                }

                var result = LayoutBuilder.Create(schemaSource, outputPath, templatePath, headingText);
                var dto = new CreateResultDto(
                    result.OutputPath,
                    result.ReportName,
                    result.ReportId,
                    result.Namespace,
                    result.StoreItemId,
                    result.UsedTemplate,
                    result.ReplacedExistingBcPart,
                    new ValidationSummaryDto(
                        result.QuickValidation.Level,
                        result.QuickValidation.Passed,
                        result.QuickValidation.ErrorCount,
                        result.QuickValidation.WarningCount));

                return ToolResponse.Success(dto);
            }
        }, schemaSource);
    }

    [McpServerTool(Name = "refresh_xml_part")]
    [Description("Update a layout's BC dataset custom XML part to a NEW schema in place - e.g. after the AL "
                 + "report dataset changed. newSchemaSource is either an existing .docx layout (its BC "
                 + "dataset part is used) or a standalone exported schema .xml. The part's CONTENT is "
                 + "replaced but its ds:itemID (storeItemID) is KEPT, so every existing binding still links "
                 + "to the same part. Every existing control binding is then classified against the new "
                 + "schema by element name: bindings that still resolve are remapped (kept, valid, "
                 + "untouched) and counted in remappedCount; bindings that no longer resolve (a column was "
                 + "renamed or deleted) are reported as orphanedBindings and LEFT IN PLACE - this tool never "
                 + "deletes or rebinds a control itself, the caller decides via remove_control/insert_field/"
                 + "insert_label. newUnboundFields is an OLD-vs-NEW DIFF: non-label leaf columns that are "
                 + "present in the new schema, absent from the old one (i.e. genuinely added by this AL "
                 + "dataset change), and not yet bound by any control - a field that was already unbound "
                 + "before the refresh and still is is NOT reported again (see list_dataset_fields for a "
                 + "standing bound/unbound inventory instead), so refreshing against an unchanged schema "
                 + "always yields an empty list. When the new schema's "
                 + "report name/id differs from the layout's own (namespaceChanged=true), every binding's "
                 + "w:prefixMappings URI and every control's w:tag are rewritten to the new identity; the "
                 + "XPath element-name steps themselves are never rewritten - that is the 'remap where "
                 + "element names match'. STRUCTURAL GATE ONLY: like every mutating tool, the write is "
                 + "rejected only if it would introduce a NEW OpenXmlValidator structural error (leaving the "
                 + "file untouched) - orphaned bindings are quick-validation (semantic) findings, always "
                 + "ALLOWED and reported, never blocked. quickValidation therefore commonly carries "
                 + "xpath-resolves errors on a successful refresh; that is the expected orphan report "
                 + "corroborated by an independent check, not a failure.")]
    public static ToolResponse RefreshXmlPart(
        [Description("Absolute path to the .docx layout file to refresh in place.")] string layoutPath,
        [Description("Absolute path to the NEW schema: an existing .docx layout (its BC dataset part is "
                     + "used) OR a standalone exported schema .xml.")] string newSchemaSource)
    {
        return GuardRefresh(layoutPath, doc => LayoutRefresher.Refresh(doc, newSchemaSource));
    }

    [McpServerTool(Name = "preview_layout")]
    [Description("Render a mock preview of a BC Word layout: merges deterministic sample data (fills bound "
                 + "fields, expands repeating sections, fills picture placeholders) into a working copy, then "
                 + "converts that copy to PDF via Word COM or LibreOffice (whichever is available). Returns "
                 + "the merged .docx path, the PDF path (when conversion succeeded), merge stats and "
                 + "warnings, and a quick validation summary. MOCK RENDER DISCLAIMER: sample data is not real "
                 + "BC data and conversion happens outside the BC report engine, so captions, fonts, "
                 + "pagination, and other BC-specific rendering behavior may differ — this is a structural/"
                 + "binding sanity check, not a substitute for a real Business Central sandbox render. "
                 + "CONCURRENCY: calls targeting the same layoutPath (another concurrent preview, or a "
                 + "mutating edit tool - even one running in ANOTHER host process, e.g. another IDE window) "
                 + "are serialized against each other via the same in-process + cross-process lock pair every "
                 + "editing tool uses, so a preview can never observe/return a half-written or another call's "
                 + "merge output; if that lock cannot be acquired within its timeout the call fails with "
                 + "file_locked rather than racing. Always read MergedDocxPath/PdfPath back from the result "
                 + "rather than assuming fixed file names — see outputDir for how those names are derived. "
                 + "SECURITY: "
                 + "before conversion, the merged copy has every external relationship (attachedTemplate, "
                 + "externally-linked images, linked OLE objects, frame/subDocument references) stripped, so "
                 + "opening a layout cloned from an untrusted source can never make the converter reach out "
                 + "over the network/SMB; plain click-to-follow hyperlinks are unaffected. Each stripped "
                 + "relationship is reported as an external-relationship-stripped warning. RETENTION: outputs "
                 + "are NOT deleted after the call returns - they persist at the returned MergedDocxPath/"
                 + "PdfPath until either (a) the SAME layout is previewed again (its default folder is reused "
                 + "and refreshed in place), or (b) a background age-sweep removes them, which only ever "
                 + "applies to preview_layout's OWN default output root (never a caller-supplied outputDir - "
                 + "see outputDir) and only once a folder's last use is older than the retention window (7 "
                 + "days). IMPORTANT: when dataOverridesPath is used, the merged .docx (and any PDF) CONTAINS "
                 + "that real data verbatim - callers previewing sensitive data should pass their own "
                 + "outputDir and delete it themselves when done, rather than relying on the age-sweep as a "
                 + "retention control. SEEING THE RESULT: to visually inspect the rendered preview, pass "
                 + "the returned PdfPath to render_preview_pages, which returns the pages as inline PNG "
                 + "image blocks.")]
    public static ToolResponse PreviewLayout(
        [Description("Absolute path to the .docx layout file.")] string layoutPath,
        [Description("Number of sample rows to generate per repeating section. Default 3. Bounded per "
                     + "repeater/data item by an internal safeguard (default cap 100): a value above the "
                     + "cap never generates or merges more than the cap's worth of rows for any single "
                     + "repeating section, reported via a row-cap warning rather than silently ignored.")]
        int rows = 3,
        [Description("Seed for the deterministic sample-data generator. Default 12345.")] int seed = 12345,
        [Description("PDF converter to use: 'auto' (default; prefers Word, falls back to LibreOffice), "
                     + "'word', or 'libreoffice'.")] string converter = "auto",
        [Description("Optional absolute path to a real exported BC dataset XML to merge with instead of "
                     + "generated sample data. Both encodings BC produces are accepted (sniffed by root "
                     + "element): the layout's own data-store part shape (NavWordReportXmlPart) and the "
                     + "report UI's Send to > XML export (ReportDataSet) - for the export shape, each "
                     + "column carrying a decimalformatter attribute holds a raw number and is formatted "
                     + "with the export's formatRegion culture, matching what BC itself renders; an export "
                     + "from a DIFFERENT report than the layout's dataset is rejected with an error naming "
                     + "both report ids.")] string? dataOverridesPath = null,
        [Description("Optional absolute output directory for the merged .docx and preview .pdf. Defaults to "
                     + "a per-layout folder under the system temp directory, named "
                     + "'<layout-basename>-<hash>' where hash is derived from the layout's full path — so "
                     + "two DIFFERENT layouts that merely share a file name (e.g. 'C:\\appA\\Invoice.docx' vs "
                     + "'C:\\appB\\Invoice.docx') never collide, while re-previewing the SAME layout always "
                     + "reuses the same folder (merged.docx/preview.pdf are refreshed in place). When "
                     + "outputDir IS supplied, that same hash is instead embedded in the FILE names "
                     + "('merged-<hash>.docx' / 'preview-<hash>.pdf') so two different layouts previewed "
                     + "into the SAME caller-chosen folder still can't collide; files always land inside the "
                     + "folder you pass. Either way, read the actual paths back from the result rather than "
                     + "assuming fixed names. When you supply outputDir yourself, YOU own its lifetime — the "
                     + "tool never sweeps or deletes anything inside it (see RETENTION above); clean it up "
                     + "yourself if it may hold real data.")] string? outputDir = null)
    {
        return Guard(() =>
        {
            if (!File.Exists(layoutPath))
            {
                throw new FileNotFoundException("layoutPath does not point to an existing file.", layoutPath);
            }

            if (!TryParseConverterKind(converter, out var kind))
            {
                return ToolResponse.Failure(
                    "invalid_argument",
                    $"Unknown converter '{converter}'.",
                    "converter must be 'auto', 'word', or 'libreoffice' (case-insensitive); 'auto' prefers "
                    + "Word COM and falls back to LibreOffice.");
            }

            // Serializes this preview against every OTHER call targeting the SAME layout path — another
            // concurrent preview, or a mutating edit tool (both go through GuardMutate's identical
            // EditLockFor). Without this, two concurrent previews of the same layout (or of two DIFFERENT
            // layouts that happen to share a basename, before the keying fix below) could interleave their
            // merge/convert steps against shared files — including the
            // narrow window where one preview could silently return ANOTHER layout's rendered PDF with
            // ok:true. The whole merge+convert+read sequence is covered, not just the merge, because the
            // final LayoutValidator.Quick(layoutPath) read below also targets the original file and must
            // not race a concurrent edit's atomic swap of that same path.
            lock (EditLockFor(layoutPath))
            {
                // Cross-process half: acquired INSIDE the in-process lock above,
                // same order GuardMutate/GuardRead use — see CrossProcessLock's own remarks. Without this, the
                // in-process lock above only protects against ANOTHER THREAD IN THIS SAME HOST PROCESS; a
                // SECOND host process (another IDE window) previewing/editing the same layoutPath at the same
                // time is invisible to it entirely.
                using var crossLock = CrossProcessLock.TryAcquire(layoutPath, CrossProcessLock.MutatingTimeout);
                if (!crossLock.Acquired)
                {
                    return CrossProcessLockTimeoutFailure(layoutPath, CrossProcessLock.MutatingTimeout);
                }

                var resolvedOutputDir = outputDir ?? Path.Combine(
                    DefaultPreviewRoot(), DefaultPreviewOutputDirName(layoutPath));

                Directory.CreateDirectory(resolvedOutputDir);
                if (outputDir is null)
                {
                    // Freshen THIS call's own folder FIRST - before this call's own sweep pass below even
                    // runs - so a layout re-previewed regularly is never swept out from under itself (see
                    // TouchDefaultPreviewDir's own remarks for why this can't just be left to the
                    // filesystem). Doing this BEFORE the sweep call, rather than after, also narrows (though
                    // cannot fully close - see SweepStalePreviewDirs's remarks) a bounded concurrency race:
                    // touching first shrinks the window in which a DIFFERENT, concurrently-running preview
                    // call's own sweep could still list this folder as stale before this call marks it fresh.
                    TouchDefaultPreviewDir(resolvedOutputDir);
                }

                // Best-effort retention sweep: runs on EVERY call, regardless of
                // whether THIS call itself uses the default root or a caller-supplied outputDir, so an
                // explicit-outputDir caller still keeps the shared default root (accumulated by earlier
                // default-root calls) bounded. Scoped to DefaultPreviewRoot() only - see the method's own
                // remarks for why that folder is safe to sweep unconditionally and why a caller-supplied
                // outputDir is never passed in here.
                SweepStalePreviewDirs(DefaultPreviewRoot(), resolvedOutputDir);

                // An explicit outputDir is caller-chosen and may be shared across unrelated layouts (they
                // don't share a lock key merely by sharing a folder), so the merged/pdf FILE names
                // themselves are also keyed by this layout's hash there. The DEFAULT folder above is
                // already unique per layout (its own name embeds the hash), so plain "merged.docx"/
                // "preview.pdf" inside it is unambiguous — keeping the common case's paths short and
                // stable across repeated previews of the same layout.
                var hash = PreviewPathHash(layoutPath);
                var mergedDocxPath = outputDir is null
                    ? Path.Combine(resolvedOutputDir, "merged.docx")
                    : Path.Combine(resolvedOutputDir, $"merged-{hash}.docx");
                var pdfPath = outputDir is null
                    ? Path.Combine(resolvedOutputDir, "preview.pdf")
                    : Path.Combine(resolvedOutputDir, $"preview-{hash}.pdf");

                var mergeResult = MergeEngine.Merge(layoutPath, mergedDocxPath, new MergeOptions
                {
                    Rows = rows,
                    Seed = seed,
                    DataOverridesPath = dataOverridesPath,
                    // Sever live bindings so the rendered preview shows the merged sample data and every generated
                    // row, rather than the renderer re-syncing each control to the un-populated custom XML part.
                    FlattenBindingsForRender = true,
                    // Security hardening: this merged copy IS
                    // about to be handed to a real converter (Word COM/LibreOffice) below, so strip any
                    // external relationship (attachedTemplate, linked image, linked OLE object, frame/
                    // subDocument ref) that Word would otherwise dereference the moment it opens the file -
                    // an NTLM-hash-leak/SSRF vector unrelated to macros. See ExternalRelationshipStripper for
                    // the full strip/keep rationale (plain hyperlinks are never touched).
                    StripExternalRelationships = true,
                });

                var pdfConverter = SelectConverter(kind);
                var conversion = pdfConverter.Convert(mergedDocxPath, pdfPath);
                var quick = LayoutValidator.Quick(layoutPath);

                var dto = new PreviewResultDto(
                    MergedDocxPath: Path.GetFullPath(mergedDocxPath),
                    PdfPath: conversion.Ok ? Path.GetFullPath(pdfPath) : null,
                    ConverterUsed: pdfConverter.Name,
                    ConverterAvailable: pdfConverter.IsAvailable,
                    ConversionOk: conversion.Ok,
                    ConversionError: conversion.Error,
                    Stats: new PreviewStatsDto(
                        mergeResult.Stats.FieldsFilled,
                        mergeResult.Stats.RepeatersExpanded,
                        mergeResult.Stats.RowsGenerated,
                        mergeResult.Stats.Unresolved,
                        mergeResult.Stats.PicturesFilled),
                    Warnings: mergeResult.Warnings.Select(w => new MergeWarningDto(w.Kind, w.Message, w.Location)).ToList(),
                    QuickValidation: new ValidationSummaryDto("quick", quick.Passed, quick.ErrorCount, quick.WarningCount),
                    Disclaimer: PreviewDisclaimer);

                return ToolResponse.Success(dto);
            }
        }, layoutPath);
    }

    // ---- preview_layout helpers ----

    /// <summary>Fixed "mock render" banner returned in every <c>preview_layout</c> response.</summary>
    private const string PreviewDisclaimer =
        "MOCK RENDER: this preview is built from deterministic sample data (not real Business Central data) "
        + "and converted to PDF by an offline converter (Word COM or LibreOffice), not the BC report engine. "
        + "Expect fidelity gaps versus a real BC render — label/caption text, fonts, exact pagination, and "
        + "other BC-specific rendering behavior can differ. Use this to sanity-check structure and bindings "
        + "only; final sign-off requires rendering the layout in a genuine Business Central sandbox. "
        // The "7 days" here is prose, not a live reference to PreviewRetentionWindow below (a const string
        // cannot embed a TimeSpan) - keep the two in sync by hand if the window ever changes.
        + "RETENTION: MergedDocxPath/PdfPath are NOT deleted when this call returns — they persist until "
        + "re-previewed (same layout) or, for the tool's own default output folder only, until an age-sweep "
        + "removes them after 7 days unused; a caller-supplied outputDir is never swept. If dataOverridesPath "
        + "was used, these files contain that REAL data — clean up sensitive-data previews yourself rather "
        + "than relying on the sweep.";

    /// <summary>
    /// Retention window for <c>preview_layout</c>'s own default output root: a
    /// merged .docx/PDF pair can silently retain a caller's REAL business data forever (when
    /// <c>dataOverridesPath</c> was used) since nothing previously deleted it. Age, not call count, is the
    /// sweep signal, so a rarely-previewed-but-still-live layout is never yanked out from under an
    /// infrequent caller between visits - 7 days comfortably covers "come back next week and look at this
    /// again" while still bounding the unlimited accumulation the finding flagged. A private constant (not
    /// configurable) - this is opportunistic housekeeping on the tool's OWN scratch space, not a contract a
    /// caller could ever legitimately code against.
    /// </summary>
    private static readonly TimeSpan PreviewRetentionWindow = TimeSpan.FromDays(7);

    /// <summary>
    /// The fixed folder <c>preview_layout</c> uses for every DEFAULT (no explicit <c>outputDir</c>) output:
    /// <c>%TEMP%\bc-word-layout-mcp</c>, one level above each layout's own
    /// <see cref="DefaultPreviewOutputDirName"/> subfolder. This root is exclusively owned by this tool - no
    /// other tool/temp usage in this codebase writes under a "bc-word-layout-mcp" subfolder of
    /// <see cref="Path.GetTempPath"/>: edit working copies, LibreOffice output/profile directories, and the
    /// full-validation dry-run merge all use their own GUID-named siblings directly under the temp root
    /// instead (<c>bcwl-edit-*</c>/<c>bcwl-lo-out-*</c>/<c>bcwl-lo-profile-*</c>/<c>bcwl-full-validate-*</c> -
    /// never this subfolder); the one other tool-generated name that looks similar, <c>.bcwl-stage-*</c>, is
    /// staged in the LAYOUT's OWN directory (see <c>ToolGuards</c>' commit/rename step), not under the temp
    /// root at all, so it never appears here either. This exclusive ownership is what makes it safe for
    /// <see cref="SweepStalePreviewDirs"/> to enumerate and delete matching immediate children unconditionally
    /// (subject to the reparse-point guard, the name-shape restriction, the age check, and the this-call
    /// exclusion - see that method's own remarks) without risking a directory anything else created or
    /// depends on.
    /// </summary>
    private static string DefaultPreviewRoot() => Path.Combine(Path.GetTempPath(), "bc-word-layout-mcp");

    /// <summary>
    /// True when <paramref name="dirName"/> has the exact shape <see cref="DefaultPreviewOutputDirName"/>
    /// itself always produces: some non-empty basename, then a single '-', then exactly the 12 lowercase hex
    /// characters <see cref="PreviewPathHash"/> always emits - and nothing after. <see cref="SweepStalePreviewDirs"/>
    /// uses this to restrict delete CANDIDATES to folder names the tool itself could plausibly have created,
    /// closing a gap review found: excluding only THIS call's own directory (<c>thisCallDir</c>) protects
    /// that one folder for THIS call only - a caller who points <c>outputDir</c> at some OTHER, arbitrarily
    /// named folder directly inside the default root would previously have had it swept by a LATER, unrelated
    /// preview call the moment it looked stale, contradicting the documented "a caller-supplied outputDir is
    /// never swept" promise. With this filter, an arbitrarily-named folder is never even a sweep candidate,
    /// regardless of age - only names shaped exactly like the tool's own default output folders are.
    /// </summary>
    private static bool LooksLikeToolOwnedPreviewDirName(string dirName)
    {
        const int hashLength = 12;
        // Shortest possible tool-owned name is a single-character basename + '-' + the 12-char hash.
        if (dirName.Length < hashLength + 2 || dirName[dirName.Length - hashLength - 1] != '-')
        {
            return false;
        }

        for (var i = dirName.Length - hashLength; i < dirName.Length; i++)
        {
            var c = dirName[i];
            if (c is not ((>= '0' and <= '9') or (>= 'a' and <= 'f')))
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>
    /// Best-effort age-based cleanup of <paramref name="defaultRoot"/>: deletes
    /// every immediate SUBDIRECTORY whose name looks tool-owned (<see cref="LooksLikeToolOwnedPreviewDirName"/>)
    /// AND whose <see cref="Directory.GetLastWriteTimeUtc(string)"/> is older than
    /// <see cref="PreviewRetentionWindow"/>, except <paramref name="thisCallDir"/> — the directory THIS call
    /// is about to (re)use, which must never be swept out from under the very call that is writing into it,
    /// even if it happens to look stale (e.g. re-previewing a layout for the first time in months reuses its
    /// existing, long-untouched folder). Runs on EVERY <c>preview_layout</c> call - including ones that
    /// themselves write to a caller-supplied <c>outputDir</c> elsewhere - since the shared default root
    /// still benefits from being swept regardless of where THIS call's own output lands; the cost is one
    /// cheap directory listing plus, at most, deleting a handful of genuinely stale folders.
    /// <para>
    /// SCOPE: only ever touches <paramref name="defaultRoot"/> itself and its immediate children — see
    /// <see cref="DefaultPreviewRoot"/> for why that folder is exclusively tool-owned. A caller-supplied
    /// <c>outputDir</c> is never passed to this method and therefore can never be examined, let alone
    /// deleted - the caller owns that directory's entire lifecycle. Review also flagged the narrower case of
    /// a caller-supplied <c>outputDir</c> that happens to sit DIRECTLY INSIDE the default root itself: the
    /// <paramref name="thisCallDir"/> exclusion above only protects it for the call that is USING it right
    /// now, not from a later, unrelated call's sweep - so every candidate is additionally required to match
    /// <see cref="LooksLikeToolOwnedPreviewDirName"/>'s exact name shape before it is even considered,
    /// regardless of age. An arbitrarily-named caller folder never matches that shape and is therefore never
    /// a candidate at all, making the "never swept" promise hold even when such a folder happens to live
    /// inside this root.
    /// </para>
    /// <para>
    /// TEMP-SQUAT HARDENING: if <paramref name="defaultRoot"/> itself is a reparse point (junction/symlink -
    /// e.g. planted by another process, or left over from an unrelated tool that squatted the same temp
    /// path), <see cref="Directory.EnumerateDirectories(string)"/> would transparently walk and this method
    /// would then delete children of whatever the junction actually points AT, not a genuine tool-owned
    /// folder — so the reparse-point check below skips the ENTIRE sweep rather than risk deleting something
    /// this tool never created.
    /// </para>
    /// <para>
    /// BEST-EFFORT: every failure is swallowed, per directory and for the initial listing/attribute check
    /// itself. Another process may hold a file open inside a stale folder (an antivirus scan, a user with the
    /// PDF open in a viewer, a concurrent preview of a DIFFERENT layout under the race described next), the
    /// folder may have already been removed by another sweep running concurrently, or the root itself may be
    /// temporarily inaccessible — none of that may ever fail the <c>preview_layout</c> call this cleanup
    /// merely piggybacks on; a directory that fails to delete is simply left for a later sweep to retry.
    /// </para>
    /// <para>
    /// ACCEPTED RESIDUAL RACE: two concurrent <c>preview_layout</c> calls for DIFFERENT layouts (different
    /// <c>layoutPath</c> edit-lock keys, so genuinely concurrent - see <c>PreviewLayout</c>'s own locking
    /// remarks) can still, in a narrow window, have one call's sweep delete the OTHER call's folder: if call
    /// B is reviving its own &gt;7-day-stale default folder, <c>Directory.CreateDirectory</c> is a no-op on an
    /// existing folder (does not bump its timestamp) - so between that no-op and B's own
    /// <c>TouchDefaultPreviewDir</c> call, a concurrently-running call A's sweep could still list B's folder
    /// as stale (A's own <paramref name="thisCallDir"/> exclusion only ever protects A's OWN folder, never
    /// B's) and delete it. <c>PreviewLayout</c> now touches its own folder BEFORE running its own sweep pass
    /// specifically to shrink this window, but cannot close it outright — the residual sliver where A's sweep
    /// runs after B's <c>CreateDirectory</c> no-op but before B's touch remains possible. Accepted as-is: the
    /// race is narrow, self-healing (the next preview of the affected layout simply regenerates its output),
    /// and can only ever destroy regenerable preview output — never a caller-supplied <c>outputDir</c> (out of
    /// scope for this sweep entirely) and never anything this tool didn't itself create.
    /// </para>
    /// </summary>
    private static void SweepStalePreviewDirs(string defaultRoot, string thisCallDir)
    {
        List<string> candidates;
        try
        {
            if (!Directory.Exists(defaultRoot))
            {
                return;
            }

            // Temp-squat hardening (see remarks above): a reparse point here means EnumerateDirectories
            // would walk someone else's real folder tree, not this tool's own - skip the sweep entirely
            // rather than risk deleting through it. File.GetAttributes works for directories too and its own
            // IO errors are caught by the same catch below, alongside the listing call's.
            if ((File.GetAttributes(defaultRoot) & FileAttributes.ReparsePoint) != 0)
            {
                return;
            }

            candidates = Directory.EnumerateDirectories(defaultRoot).ToList();
        }
        catch
        {
            return; // Best-effort: an inaccessible/unreadable root must never fail the preview.
        }

        var thisCallFullPath = Path.GetFullPath(thisCallDir);
        var cutoffUtc = DateTime.UtcNow - PreviewRetentionWindow;

        foreach (var candidate in candidates)
        {
            // Name-shape restriction (follow-up): only ever consider folders shaped exactly like
            // the tool's own DefaultPreviewOutputDirName output - see LooksLikeToolOwnedPreviewDirName and
            // this method's own remarks for why an arbitrarily-named caller folder must never even reach the
            // age check below, regardless of where it happens to sit.
            if (!LooksLikeToolOwnedPreviewDirName(Path.GetFileName(candidate)))
            {
                continue;
            }

            if (string.Equals(Path.GetFullPath(candidate), thisCallFullPath, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            try
            {
                if (Directory.GetLastWriteTimeUtc(candidate) < cutoffUtc)
                {
                    Directory.Delete(candidate, recursive: true);
                }
            }
            catch
            {
                // Best-effort: leave this one candidate for a later sweep rather than failing the preview.
            }
        }
    }

    /// <summary>
    /// Stamps <paramref name="previewDir"/>'s own <see cref="Directory.LastWriteTimeUtc"/> to now, so a
    /// layout that gets re-previewed regularly is never treated as stale by
    /// <see cref="SweepStalePreviewDirs"/>. This is NOT redundant with the filesystem's own behavior: a
    /// re-preview of the SAME layout reuses this exact folder and overwrites its two files
    /// (<c>merged.docx</c>/<c>preview.pdf</c>) IN PLACE via <see cref="File.Copy(string, string, bool)"/> and
    /// equivalent converter writes — which updates each FILE's own timestamp but, confirmed empirically on
    /// NTFS, does NOT bump the PARENT DIRECTORY's own <c>LastWriteTime</c> (that only moves when an entry is
    /// added/removed/renamed, not when an existing entry's content is merely overwritten). Without this
    /// explicit stamp, a layout previewed daily forever would still eventually cross
    /// <see cref="PreviewRetentionWindow"/> and be swept alive by <see cref="SweepStalePreviewDirs"/>, purely
    /// because its folder's own timestamp never moved past its original creation time — silently destroying
    /// live, actively-used output. Only ever called for the tool's own default-root folder (never a
    /// caller-supplied <c>outputDir</c> — see <c>PreviewLayout</c>'s call site) and swallows any failure: a
    /// missed freshness stamp only makes THIS folder look one call staler than it is, never a reason to fail
    /// the preview itself.
    /// </summary>
    private static void TouchDefaultPreviewDir(string previewDir)
    {
        try
        {
            Directory.SetLastWriteTimeUtc(previewDir, DateTime.UtcNow);
        }
        catch
        {
            // Best-effort - see remarks above.
        }
    }

    private static bool TryParseConverterKind(string converter, out PdfConverterKind kind)
    {
        switch (converter.Trim().ToLowerInvariant())
        {
            case "auto":
                kind = PdfConverterKind.Auto;
                return true;
            case "word":
                kind = PdfConverterKind.Word;
                return true;
            case "libreoffice":
                kind = PdfConverterKind.LibreOffice;
                return true;
            default:
                kind = default;
                return false;
        }
    }

    /// <summary>Replaces characters invalid in a path segment with '_', so an arbitrary layout file name is
    /// always safe to use as the default preview output folder name. '.'/'..' are NOT in
    /// <see cref="Path.GetInvalidFileNameChars"/>, so e.g. a layout named "...docx" would otherwise sanitize
    /// to the untouched special segment ".." and silently escape up one directory level (landing the
    /// default output directly in the shared temp root instead of a per-layout subfolder) — guarded
    /// explicitly below rather than relying on the invalid-char replacement to catch it.</summary>
    private static string SanitizeForPath(string name)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var chars = name.Select(c => invalid.Contains(c) ? '_' : c).ToArray();
        var sanitized = new string(chars);
        return string.IsNullOrWhiteSpace(sanitized) || sanitized is "." or ".." ? "layout" : sanitized;
    }

    /// <summary>
    /// First 12 lowercase hex characters of SHA-256 over the layout's normalized full path, used to key
    /// <c>preview_layout</c>'s default output folder (<see cref="DefaultPreviewOutputDirName"/>) and, when
    /// the caller supplies an explicit <c>outputDir</c>, the merged/pdf FILE names within it (see
    /// <c>PreviewLayout</c>). The path is case-folded via <see cref="string.ToUpperInvariant"/> before
    /// hashing (Windows paths are case-insensitive; hashing the raw bytes would otherwise treat
    /// <c>C:\A.docx</c> and <c>c:\a.docx</c> as different files) - the same normalization
    /// <see cref="ToolGuards.EditLockFor"/> relies on via its <see cref="StringComparer.OrdinalIgnoreCase"/>-keyed
    /// lock dictionary, so any two path spellings that share an edit lock also share this hash.
    /// <para>
    /// Full SHA-256 (not just <see cref="string.GetHashCode"/>) is used because .NET's own string hash
    /// is randomized per process run - it would defeat the whole point of a STABLE per-layout folder name
    /// (same layout re-previewed later, possibly by a different host process, must resolve to the same
    /// directory). 12 hex chars (48 bits) is far more than enough to make an accidental collision between
    /// two DIFFERENT layout paths that happen to share a basename effectively impossible, while keeping
    /// the folder/file name short and human-scannable.
    /// </para>
    /// Internal rather than private so it is directly unit-testable (see the
    /// <c>InternalsVisibleTo</c> grant in BcWordLayout.McpHost.csproj).
    /// </summary>
    internal static string PreviewPathHash(string layoutPath)
    {
        var normalized = Path.GetFullPath(layoutPath).ToUpperInvariant();
        var hash = System.Security.Cryptography.SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(normalized));
        return Convert.ToHexString(hash)[..12].ToLowerInvariant();
    }

    /// <summary>
    /// Default <c>preview_layout</c> output folder NAME (a single path segment - the caller still combines
    /// it under <see cref="Path.GetTempPath"/> and a fixed "bc-word-layout-mcp" root):
    /// <c>&lt;sanitized-basename&gt;-&lt;hash&gt;</c>, where hash is <see cref="PreviewPathHash"/>.
    /// <para>
    /// Keying by the FULL normalized path (not just the basename <see cref="SanitizeForPath"/> sanitizes)
    /// is what makes two DIFFERENT layouts that merely share a file name - e.g.
    /// <c>C:\appA\SalesInvoice.docx</c> vs. <c>C:\appB\SalesInvoice.docx</c> - resolve to different
    /// folders, while the SAME layout re-previewed - regardless of
    /// case or of a relative-vs-absolute spelling, anything <see cref="Path.GetFullPath(string)"/>
    /// normalizes away - always resolves to the SAME folder, so its <c>merged.docx</c>/<c>preview.pdf</c>
    /// are simply refreshed in place rather than accumulating stale copies under an ever-growing set of
    /// directories. The basename prefix is kept purely so the folder stays human-recognizable in a
    /// directory listing; only the hash suffix is load-bearing for uniqueness.
    /// </para>
    /// </summary>
    internal static string DefaultPreviewOutputDirName(string layoutPath) =>
        $"{SanitizeForPath(Path.GetFileNameWithoutExtension(layoutPath))}-{PreviewPathHash(layoutPath)}";

    // ---- render_preview_pages ----

    /// <summary>
    /// The ONE tool in this server whose result is not purely the JSON envelope: it returns a
    /// <see cref="CallToolResult"/> directly (the SDK passes that through verbatim — every other return
    /// type gets serialized into a single text block) so it can carry MCP IMAGE content blocks alongside
    /// the usual <see cref="ToolResponse"/> JSON, letting a calling agent literally look at the rendered
    /// preview instead of only receiving a file path its client may or may not be able to open.
    /// <para>
    /// No lock pair is taken: this reads ONE scratch file (the preview PDF) into memory in a single
    /// <see cref="File.ReadAllBytes"/> and never touches the layout itself — there is no layoutPath here
    /// to key a lock on, and the worst a concurrent re-preview of the same layout can do is make that
    /// read fail (sharing violation → clean failure envelope) or hand us the OLD complete PDF (converters
    /// write to the final path only on success). A torn read cannot produce a half-corrupt image: PDFium
    /// operates on the in-memory snapshot and rejects a truncated document as a whole.
    /// </para>
    /// </summary>
    [McpServerTool(Name = "render_preview_pages")]
    [Description("Render pages of a preview PDF (the PdfPath returned by preview_layout) as PNG images "
                 + "returned INLINE as MCP image content blocks, so the calling agent can visually inspect "
                 + "the preview without needing file-system or PDF support of its own. The first content "
                 + "block is the usual JSON envelope (page count, dimensions, truncation flag); each "
                 + "rendered page follows as one image block, in page order. COST: every page image adds "
                 + "on the order of a thousand-plus tokens to the response — request only the pages you "
                 + "need (default 3, hard cap 10 per call; page through longer documents with firstPage). "
                 + "The same MOCK RENDER disclaimer as preview_layout applies: this shows the offline-"
                 + "converted preview, not a genuine Business Central render.")]
    public static CallToolResult RenderPreviewPages(
        [Description("Absolute path to the PDF to render — typically the PdfPath returned by a preceding "
                     + "preview_layout call.")] string pdfPath,
        [Description("1-based page number to start rendering from. Default 1.")] int firstPage = 1,
        [Description("How many pages to return, starting at firstPage. Default 3, hard cap 10 per call — "
                     + "the JSON envelope's 'truncated' flag tells you when the document has more pages "
                     + "than were returned.")] int maxPages = 3,
        [Description("Render resolution in DPI, clamped to 36–300. Default 120 (an A4 page comes out "
                     + "roughly 992×1403 px — readable without wasting tokens on print-resolution "
                     + "detail).")] int dpi = 120)
    {
        // File-existence gets its own code (and the shared file_not_found hint style Guard uses) rather
        // than the rasterizer's generic error string, so an agent that cached a stale PdfPath (e.g. the
        // preview was re-run or age-swept) is told exactly how to recover: preview again, then re-render.
        if (!File.Exists(pdfPath))
        {
            return EnvelopeOnly(ToolResponse.Failure(
                "file_not_found",
                $"pdfPath does not point to an existing file. Path checked: '{pdfPath}'.",
                "Pass the PdfPath returned by a preview_layout call. Preview outputs are refreshed in "
                + "place on re-preview and age-swept after long disuse, so a stale path from an old call "
                + "may be gone — run preview_layout again and use the PdfPath it returns."));
        }

        var result = PdfRasterizer.Rasterize(pdfPath, new PdfRasterizeOptions
        {
            FirstPage = firstPage,
            MaxPages = maxPages,
            Dpi = dpi,
        });

        if (!result.Ok)
        {
            return EnvelopeOnly(ToolResponse.Failure(
                "rasterize_failed",
                result.Error ?? "PDF rasterization failed.",
                "Confirm pdfPath is the PdfPath from a successful preview_layout call (conversionOk true) "
                + "and that firstPage is within the document (see the message for the page count when "
                + "available); if the file was replaced by a concurrent re-preview, run preview_layout "
                + "again and retry with its fresh PdfPath."));
        }

        var dto = new PreviewPagesResultDto(
            PdfPath: Path.GetFullPath(pdfPath),
            PageCount: result.PageCount ?? result.Pages.Count,
            FirstPage: firstPage,
            PagesRendered: result.Pages.Count,
            EffectiveDpi: result.EffectiveDpi,
            Truncated: result.Truncated,
            Pages: result.Pages
                .Select(p => new PreviewPageDto(p.PageNumber, p.WidthPx, p.HeightPx, p.PngBytes.Length))
                .ToList());

        var content = new List<ContentBlock>(1 + result.Pages.Count)
        {
            SerializeEnvelope(ToolResponse.Success(dto)),
        };
        foreach (var page in result.Pages)
        {
            content.Add(ImageContentBlock.FromBytes(page.PngBytes, "image/png"));
        }

        return new CallToolResult { Content = content };
    }

    /// <summary>A failure <see cref="CallToolResult"/> carrying only the JSON envelope, no image blocks.</summary>
    private static CallToolResult EnvelopeOnly(ToolResponse envelope) =>
        new() { Content = [SerializeEnvelope(envelope)] };

    /// <summary>
    /// Serializes a <see cref="ToolResponse"/> to a text content block with the SAME serializer options
    /// the SDK itself applies to every other tool's returned envelope (<see cref="McpJsonUtilities.DefaultOptions"/>
    /// is documented as the default for "other types" marshaling) — so this tool's JSON half is
    /// byte-for-byte the shape agents already parse from its siblings.
    /// </summary>
    private static TextContentBlock SerializeEnvelope(ToolResponse envelope) =>
        new() { Text = JsonSerializer.Serialize(envelope, McpJsonUtilities.DefaultOptions) };
}
