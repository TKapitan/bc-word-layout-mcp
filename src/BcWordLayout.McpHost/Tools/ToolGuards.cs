using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace BcWordLayout.McpHost.Tools;

/// <summary>
/// Shared infrastructure for every <c>[McpServerToolType]</c> tool-family class in this namespace
/// (<see cref="ReadTools"/>, <see cref="EditTools"/>, <see cref="TableTools"/>,
/// <see cref="LifecycleTools"/>) — none of these methods are themselves MCP tools (this class carries no
/// <c>[McpServerToolType]</c>/<c>[McpServerTool]</c> attributes, so <c>WithToolsFromAssembly()</c> never
/// discovers it). It holds: the exception-to-envelope translator (<see cref="Guard"/>) and its three hint
/// maps (<see cref="InvalidArgumentHint"/>, <see cref="NotFoundHintFor"/>, plus the inline <c>file_locked</c>
/// hints — see <see cref="Guard"/>'s own <c>IOException</c> branch); the mutating-tool open/validate/save-
/// or-reject choreography (<see cref="GuardMutate{TResult}"/>) and its four per-family thin instantiations
/// (<see cref="GuardEdit"/>, <see cref="GuardCellEdit"/>, <see cref="GuardTableEdit"/>,
/// <see cref="GuardRefresh"/>); the READ-tool counterpart (<see cref="GuardRead"/>); the per-path edit-lock
/// dictionary (<see cref="EditLockFor"/>) and its cross-process counterpart (<see cref="CrossProcessLock"/>);
/// the flat-parameter <see cref="Location"/> builder (<see cref="BuildLocation"/>); and shared response DTO-
/// mapping helpers (<see cref="ToTableDto"/>, <see cref="ToDataItemDto"/>, etc.). Splitting the tools
/// themselves by family only works because every family shares this one place
/// for the plumbing they all depend on identically.
/// </summary>
internal static class ToolGuards
{
    // ---- mutating-tool edit safety (open read-write copy, mutate, validate BEFORE save, atomic commit) ----

    /// <summary>
    /// Per-path locks serializing concurrent edits targeting the same layout file (keyed by the normalized
    /// full path, case-insensitively — <see cref="StringComparer.OrdinalIgnoreCase"/> is the dictionary's
    /// own comparer, so lookups need no manual case-folding). Two <see cref="GuardEdit"/> calls against
    /// DIFFERENT files never contend; two against the SAME file are serialized so their working-copy /
    /// stage-and-swap sequences can never race each other (e.g. both reading the same "pre-existing ids"
    /// snapshot and generating colliding <c>w:id</c>s, or one's atomic swap clobbering the other's). Entries
    /// are never evicted — for a developer-facing tool touching a bounded number of distinct layout paths
    /// per process lifetime, that is a handful of small object references, not a real leak.
    /// <para>
    /// This is the IN-PROCESS half of this server's locking story only — it serializes threads within ONE
    /// host process. Two SEPARATE host processes (e.g. two IDE windows) touching the same path are invisible
    /// to each other here; that cross-process half is <see cref="CrossProcessLock"/>, always acquired
    /// INSIDE this lock, never instead of it — see <see cref="CrossProcessLock"/>'s
    /// own remarks for the full acquire-order rationale.
    /// </para>
    /// <para>
    /// CORRECTNESS INVARIANT: this dictionary MUST remain a single instance shared by every mutating tool
    /// across every tool-family class (<see cref="EditTools"/>, <see cref="TableTools"/>,
    /// <see cref="LifecycleTools"/>'s <c>refresh_xml_part</c>/<c>preview_layout</c>) — that is what makes
    /// "two mutating tools on the same path serialize against each other" true regardless of which family
    /// either call belongs to. It lives here, in the one file every family's mutating tools funnel through
    /// via <see cref="EditLockFor"/>/<see cref="GuardMutate{TResult}"/>, precisely so no future split could
    /// accidentally give a different tool family its own separate dictionary and silently reintroduce the
    /// cross-family race this design prevents. <see cref="ReadTools"/>'s READ-ONLY tools now take the SAME
    /// dictionary too, via <see cref="GuardRead"/> — see its own remarks for why
    /// a read must coordinate with the mutating tools' atomic-rename commit.
    /// </para>
    /// </summary>
    private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, object> EditLocks =
        new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Internal rather than private so a test can grab the SAME lock object a real tool call would (see
    /// <c>PreviewLayout_serializes_against_a_concurrent_holder_of_the_same_layouts_edit_lock</c> in
    /// BcWordLayout.Tests) and assert the tool call blocks on it - deterministically, since holding a
    /// <c>Monitor</c> lock and observing a second thread stall on <c>Monitor.Enter</c> requires no timing
    /// assumptions about how long the guarded work itself takes.
    /// </summary>
    internal static object EditLockFor(string layoutPath) =>
        EditLocks.GetOrAdd(Path.GetFullPath(layoutPath), static _ => new object());

    /// <summary>
    /// Shared choreography for every mutating tool (<c>insert_field</c>, <c>insert_label</c>,
    /// <c>insert_repeater_table</c>, <c>remove_control</c>, <c>refresh_xml_part</c>): copy
    /// <paramref name="layoutPath"/> to a private temp working copy, open THAT copy read-write
    /// (<see cref="OpenSettings.AutoSave"/> off), apply <paramref name="mutate"/>, then run
    /// <see cref="OpenXmlValidator"/> both BEFORE and after the edit so only NEWLY introduced structural
    /// errors can reject it (see remarks). Only when clean is the working copy saved, staged into
    /// <paramref name="layoutPath"/>'s own directory, and atomically swapped into place via
    /// <see cref="File.Move(string, string, bool)"/> — the ORIGINAL file is never opened in write mode at
    /// all, and the final swap is a same-volume rename rather than a truncate-and-stream copy, so a
    /// crash/IO error/AV lock mid-write can never leave <paramref name="layoutPath"/> partially written (see
    /// remarks). Concurrent edits of the same path are serialized via <see cref="EditLockFor"/>. Exceptions
    /// from <paramref name="mutate"/> (bad dataset path, bad location, an unresolvable location, an unknown
    /// control id, a missing/invalid new schema source, …) flow up through the same
    /// <see cref="Guard(Func{ToolResponse}, string)"/> every other tool uses.
    /// <para>
    /// Structural guarantee (generalized in): this method is also the ONE place that
    /// computes a post-edit <see cref="LayoutValidator.Quick"/> summary (on the saved working copy, right
    /// before commit) AND the ONE place <paramref name="onSuccess"/> is invoked to build the tool's own
    /// success <see cref="ToolResponse"/> — every mutating tool's own DTO construction (<see cref="EditResultDto"/>,
    /// <see cref="RefreshResultDto"/>) is therefore structurally incapable of shipping without a populated
    /// <c>QuickValidation</c>; this is not a convention an individual tool could forget to follow. The
    /// <paramref name="mutate"/> result type is generic (<typeparamref name="TResult"/>) precisely because
    /// different mutating tools return different result shapes — <see cref="EditResult"/> for the four
    /// content-control edits, <see cref="RefreshResult"/> for <c>refresh_xml_part</c> (which additionally
    /// computes ITS OWN <see cref="RefreshResult.QuickValidation"/> internally, for Domain-level callers
    /// that never go through this tool layer at all — <paramref name="onSuccess"/> is free to reuse that
    /// one instead of the <see cref="ValidationResult"/> passed alongside <typeparamref name="TResult"/>,
    /// as <c>refresh_xml_part</c>'s own callback does, since both reflect the identical post-mutate state).
    /// </para>
    /// </summary>
    /// <remarks>
    /// Two things this method does NOT do, and why:
    /// <list type="bullet">
    /// <item>It does not open <paramref name="layoutPath"/> itself read-write and just skip <c>Save()</c> on
    /// failure. An earlier version did exactly that (AutoSave off, validate, save-or-skip); a dedicated test
    /// caught it corrupting the file anyway — merely opening an OPC package read-write and reading so much
    /// as one part's raw stream (which <c>SchemaProvider</c> does) can make the underlying zip container get
    /// repackaged/recompressed on <c>Dispose()</c> regardless of AutoSave or whether any typed part was ever
    /// <c>Save()</c>'d, which showed up as a several-byte drift on a failed edit with no <c>Save()</c> ever
    /// called. Never opening the real file for writing, and only ever touching it via one atomic rename,
    /// sidesteps that entirely rather than depending on the exact repackaging trigger. Likewise the final
    /// commit stages into <paramref name="layoutPath"/>'s own directory rather than leaving the validated
    /// result in <see cref="Path.GetTempPath"/> and copying it over — <see cref="Path.GetTempPath"/> is
    /// frequently a different volume, where <see cref="File.Move(string, string, bool)"/> silently degrades
    /// to a non-atomic copy+delete.
    /// </item>
    /// <item>It does not reject on any structural error present after the edit, only on ones ABSENT before
    /// it. A layout that already has pre-existing OpenXML errors of its own must stay editable via these
    /// tools rather than becoming permanently stuck, and any such pre-existing errors must never be
    /// misreported as caused by this edit. For <c>refresh_xml_part</c> specifically, this is also what makes
    /// it a STRUCTURAL gate only: an orphaned binding is a <see cref="LayoutValidator.Quick"/> (semantic)
    /// finding, never an <see cref="OpenXmlValidator"/> (structural) one, so it can never trigger the
    /// rejection below — it is always allowed through and reported, exactly as designed.</item>
    /// </list>
    /// </remarks>
    /// <summary>
    /// Runs a copy/move of the working-copy commit path, absorbing the TRANSIENT access-denied /
    /// sharing-violation errors an antivirus or indexer scan holds on a freshly written .docx for a few
    /// tens of milliseconds (observed as one-off "Access to the path is denied." failures under parallel
    /// load on Windows — Defender scans every new Office file). Bounded: a few short backoff attempts,
    /// then the original exception propagates unchanged, so a PERSISTENT hold (the file genuinely open in
    /// Word) still lands on <c>Guard</c>'s file_locked / io mapping exactly as before.
    /// </summary>
    private static void RetryTransientFileDenial(Action fileOperation) => TransientFileRetry.Run(fileOperation);

    private static ToolResponse GuardMutate<TResult>(
        string layoutPath, Func<WordprocessingDocument, TResult> mutate, Func<TResult, ValidationResult, ToolResponse> onSuccess)
    {
        return Guard(() =>
        {
            lock (EditLockFor(layoutPath))
            {
                if (!File.Exists(layoutPath))
                {
                    throw new FileNotFoundException("layoutPath does not point to an existing file.", layoutPath);
                }

                // Cross-process half of the edit lock: acquired INSIDE the
                // in-process lock above, never instead of it — see CrossProcessLock's own remarks for why
                // that order is deadlock-free. A SECOND host process (another IDE window) currently editing/
                // previewing/reading this SAME path is invisible to the in-process lock alone, since it lives
                // in a different process's memory entirely; this named OS mutex is what the two processes
                // actually contend on. Timing out here does not throw across the MCP boundary — it returns
                // the same structured file_locked failure the file_locked IOException branch in Guard below
                // produces for the "someone has it open in Word" case, so both flavors of "this file is busy
                // right now" land on the identical, agent-actionable error code.
                using var crossLock = CrossProcessLock.TryAcquire(layoutPath, CrossProcessLock.MutatingTimeout);
                if (!crossLock.Acquired)
                {
                    return CrossProcessLockTimeoutFailure(layoutPath, CrossProcessLock.MutatingTimeout);
                }

                // Best-effort orphan cleanup: runs on EVERY mutating call,
                // scoped to layoutPath's OWN directory only — see SweepStaleStagingFiles's own remarks.
                SweepStaleStagingFiles(layoutPath);

                var workingCopy = Path.Combine(Path.GetTempPath(), $"bcwl-edit-{Guid.NewGuid():N}.docx");
                RetryTransientFileDenial(() => File.Copy(layoutPath, workingCopy, overwrite: true));
                try
                {
                    TResult mutateResult;
                    ValidationResult quick;

                    using (var doc = WordprocessingDocument.Open(workingCopy, isEditable: true, new OpenSettings { AutoSave = false }))
                    {
                        // Baseline BEFORE the edit: see the "does NOT reject" remark above.
                        var baselineErrors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc)
                            .Select(e => e.Description)
                            .ToList();

                        // Second baseline: content controls nested inside a plain-text control. Word rejects
                        // these as a corrupt document, but OpenXmlValidator does NOT (it is a Word semantic
                        // rule, not a schema rule), so the OpenXmlValidator gate below cannot catch them - see
                        // PlainTextNestingGuard. Diffed BEFORE-vs-AFTER by description, exactly like the
                        // structural errors above, so only nestings THIS edit introduces are rejected.
                        var baselineNesting = PlainTextNestingGuard.Find(doc).Select(n => n.Describe()).ToList();

                        // Third baseline: table rows whose cell layout is inconsistent with their w:tblGrid
                        // (cells covering the wrong number of grid columns, or an empty w:tc). Like nested
                        // plain-text controls, Word treats these as a broken/corrupt table but OpenXmlValidator
                        // accepts them silently, so the structural gate below cannot catch them - see
                        // TableGridConsistencyGuard. Diffed BEFORE-vs-AFTER by STABLE IDENTITY, not description
                        // (see DiffTableGridViolations and TableGridViolation's own
                        // remarks), so only a desync THIS edit introduces is rejected (a pre-existing ragged
                        // table stays editable even across insert_column/remove_column's own gridCount change).
                        var baselineTableGrid = TableGridConsistencyGuard.Find(doc);

                        mutateResult = mutate(doc);

                        var newNesting = DiffOut(baselineNesting, PlainTextNestingGuard.Find(doc).Select(n => n.Describe()));
                        if (newNesting.Count > 0)
                        {
                            return ToolResponse.Failure(
                                "edit_would_corrupt",
                                $"This edit would place {newNesting.Count} content control(s) inside a "
                                + "plain-text content control, which Word rejects as a corrupt document; the "
                                + "file was NOT modified.",
                                $"Offending nesting: {string.Join(" | ", newNesting.Take(5))}. A plain-text "
                                + "content control (the shape BC uses for most header fields, e.g. an address "
                                + "or document-number cell) can only contain text, never another field/label "
                                + "control. The location you targeted resolves to a spot INSIDE that control - "
                                + "most often 'afterControl' targeting a cell-level field (which anchors inside "
                                + "that field's own cell) or 'atText'/'tableCell' landing in such a cell. Insert "
                                + "the new control into a DIFFERENT cell (its own column/cell via 'tableCell'), "
                                + "or at a location that is not inside an existing field control, instead.");
                        }

                        var newTableGrid = DiffTableGridViolations(baselineTableGrid, TableGridConsistencyGuard.Find(doc));
                        if (newTableGrid.Count > 0)
                        {
                            return ToolResponse.Failure(
                                "edit_would_break_table",
                                $"This edit would leave {newTableGrid.Count} table row(s) inconsistent with their "
                                + "grid (cells covering the wrong number of grid columns, or an empty cell), which "
                                + "Word renders as a broken/corrupt table; the file was NOT modified.",
                                $"Offending row(s): {string.Join(" | ", newTableGrid.Take(5))}. A table's rows must "
                                + "each cover exactly as many grid columns as its w:tblGrid declares (summing "
                                + "gridSpan), and every cell must hold content. This should not happen for any "
                                + "documented input to the table tools, so it likely indicates a bug in the tool "
                                + "rather than your call; the file on disk is unchanged. If you were targeting a "
                                + "table that uses vertical merges (vMerge), that shape is not supported yet "
                                + "(GitHub issue #9).");
                        }

                        var afterErrors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc).ToList();

                        // Match each post-edit error against one still-unmatched baseline error (by
                        // description, since positional XPaths can legitimately shift when this edit
                        // inserts/removes an unrelated sibling elsewhere in the tree); whatever is left over
                        // is genuinely NEW - introduced by this edit, not pre-existing.
                        var remainingBaseline = new List<string?>(baselineErrors);
                        var newErrors = new List<ValidationErrorInfo>();
                        foreach (var error in afterErrors)
                        {
                            var idx = remainingBaseline.IndexOf(error.Description);
                            if (idx >= 0)
                            {
                                remainingBaseline.RemoveAt(idx);
                            }
                            else
                            {
                                newErrors.Add(error);
                            }
                        }

                        // In practice this is an unreachable safety net, not a normally-exercised path:
                        // SdtFactory, LocationResolver, and LayoutRefresher are all deliberately engineered
                        // (3.2/3.4/4.1) to only ever emit/target valid OOXML, so no known route
                        // through insert_field/insert_label/remove_control/insert_repeater_table/
                        // refresh_xml_part's public surface can actually trigger it today. It is kept as a
                        // backstop against a future change breaking that invariant, and is deliberately not
                        // force-tested — there is no honest way to reach it from the outside without
                        // fabricating an internal bug to order.
                        if (newErrors.Count > 0)
                        {
                            var preview = string.Join(
                                " | ", newErrors.Take(5).Select(e => $"{e.Path?.XPath}: {e.Description}"));
                            return ToolResponse.Failure(
                                "edit_would_corrupt",
                                $"This edit would introduce {newErrors.Count} new OpenXML structural error(s); "
                                + "the file was NOT modified.",
                                $"First new error(s): {preview} - this should not happen for any documented "
                                + "input to this tool, so it likely indicates a bug in the tool itself rather "
                                + "than a mistake in your call; the file on disk is unchanged, so it is safe "
                                + "to retry with a different edit, but consider reporting this with the "
                                + "details above.");
                        }

                        // Only save parts that could plausibly have changed; Save() on an untouched,
                        // already-loaded part is cheap, so saving all three unconditionally (matching
                        // MergeEngine.Merge's own save pattern) is simpler than tracking exactly which one a
                        // given operation touched.
                        doc.MainDocumentPart!.Document!.Save();
                        foreach (var header in doc.MainDocumentPart.HeaderParts)
                        {
                            header.Header?.Save();
                        }

                        foreach (var footer in doc.MainDocumentPart.FooterParts)
                        {
                            footer.Footer?.Save();
                        }

                        // Computed exactly ONCE here, for every mutating tool alike - see this method's own
                        // "Structural guarantee" doc-comment remark above.
                        quick = LayoutValidator.Quick(doc);
                    }

                    // Commit: stage the validated result IN layoutPath's OWN DIRECTORY (guaranteeing the
                    // swap below is a same-volume rename), then atomically replace layoutPath with it. This
                    // is the one moment layoutPath actually changes, and it happens as a single filesystem
                    // rename rather than a truncate-then-stream copy, so it can never leave a partially
                    // written file behind even if the process is killed mid-operation.
                    // Path.GetDirectoryName returns null only for a bare root (e.g. "C:\") - a path-shape
                    // edge case, not a lookup failure; left as a plain InvalidOperationException
                    // (→ internal_error via Guard's generic catch), not NotFoundException.
                    var targetDir = Path.GetDirectoryName(Path.GetFullPath(layoutPath))
                        ?? throw new InvalidOperationException($"Could not determine the directory of '{layoutPath}'.");
                    var staged = Path.Combine(targetDir, $".bcwl-stage-{Guid.NewGuid():N}.docx");
                    RetryTransientFileDenial(() => File.Copy(workingCopy, staged, overwrite: true));
                    try
                    {
                        RetryTransientFileDenial(() => File.Move(staged, layoutPath, overwrite: true));
                    }
                    catch
                    {
                        if (File.Exists(staged))
                        {
                            File.Delete(staged);
                        }

                        throw;
                    }

                    return onSuccess(mutateResult, quick);
                }
                finally
                {
                    if (File.Exists(workingCopy))
                    {
                        File.Delete(workingCopy);
                    }
                }
            }
        }, layoutPath);
    }

    /// <summary>
    /// Retention window for the orphan-staging-file sweep below: a crash between
    /// <see cref="File.Copy(string, string, bool)"/> and <see cref="File.Move(string, string, bool)"/> in
    /// the commit step above is normally over in microseconds, so any genuinely orphaned staging file is,
    /// in practice, either brand new (another in-flight commit — this process or a different one entirely,
    /// targeting a DIFFERENT layout in the same directory) or has been sitting untouched for a very long
    /// time (a hard-killed process that never reached its own cleanup). One day comfortably clears any
    /// plausible in-flight window (network share latency, an AV scan holding the handle, a slow disk)
    /// while still being "soon" on human timescales, so an orphan left behind by a crash doesn't linger in
    /// a source-controlled layout's directory indefinitely.
    /// </summary>
    /// <remarks>
    /// AGE SIGNAL — <see cref="File.GetCreationTimeUtc(string)"/>, deliberately NOT
    /// <see cref="File.GetLastWriteTimeUtc(string)"/> (a bug found and fixed during review of this exact
    /// method): <see cref="File.Copy(string, string, bool)"/> copies the SOURCE's last-write time onto the
    /// new file but resets its creation time to "now" (confirmed empirically on Windows/NTFS). The staged
    /// file here is itself produced by a copy (of a working copy, itself copied from <c>layoutPath</c>) — if
    /// <c>layoutPath</c> is an old, already-committed file (the common case: a real corpus/checked-in
    /// layout), its last-write time can trivially be more than a day old, making a staging file created THIS
    /// INSTANT already look "stale" by last-write time alone — exactly the failure mode that let a
    /// concurrent sweep delete a live, mid-commit staging file out from under an in-progress edit. Creation
    /// time reliably reflects how long THIS staging file has actually existed, regardless of its content's
    /// own history.
    /// </remarks>
    private static readonly TimeSpan StageFileRetentionWindow = TimeSpan.FromDays(1);

    /// <summary>
    /// Name shapes swept by <see cref="SweepStaleStagingFiles"/>: this class's own <c>.bcwl-stage-*.docx</c>
    /// commit-step staging file, AND <c>BcWordLayout.Merge.MergeEngine</c>'s <c>.bcwl-merge-stage-*.docx</c>
    /// merge-commit staging file (interaction). The two live in DIFFERENT
    /// directories in the common case — this one in a layout's own directory, the merge one in whatever
    /// directory a <c>preview_layout</c>/merge <c>outputDir</c> points at — so this sweep, which only ever
    /// looks at a layout's own directory, is a complementary/opportunistic pass for the merge shape (the
    /// primary defense for it is <c>MergeEngine.SweepStaleMergeStagingFiles</c>, scoped to the merge
    /// output's own directory instead). Swept here too anyway, in case a caller happens to point a
    /// merge/preview <c>outputDir</c> AT a layout's own directory, where nothing else would ever revisit it.
    /// </summary>
    private static readonly string[] StagingFileGlobs = { ".bcwl-stage-*.docx", ".bcwl-merge-stage-*.docx" };

    /// <summary>
    /// Best-effort age-based cleanup of orphaned staging files (see
    /// <see cref="StagingFileGlobs"/> for the exact name shapes): the commit step above creates
    /// <c>.bcwl-stage-*.docx</c> IN <paramref name="layoutPath"/>'s OWN directory (to guarantee the final
    /// swap is a same-volume rename) and normally deletes it again within the same call — either by the
    /// successful <see cref="File.Move(string, string, bool)"/> consuming it, or by the <c>catch</c>
    /// block's explicit delete on failure. A process kill in the narrow window between the copy and the
    /// move (or the move and its own catch) skips both of those, leaving the file behind next to
    /// source-controlled layouts, where it shows up in <c>git status</c> and never self-heals — nothing
    /// else ever revisits it. This sweep runs on EVERY mutating call (before that call creates its own
    /// staged file), scoped ONLY to <paramref name="layoutPath"/>'s own directory — never a caller's whole
    /// workspace, and never anything outside the one folder this call is about to write into anyway.
    /// </summary>
    /// <remarks>
    /// <para>
    /// SAFE AGAINST A LIVE CONCURRENT COMMIT: only files older than <see cref="StageFileRetentionWindow"/>
    /// are removed — comfortably longer than any plausible in-flight commit (see that field's own
    /// remarks) — so a staging file another call (this process, serialized behind the very
    /// <see cref="EditLockFor"/>/<see cref="CrossProcessLock"/> pair guarding this one, or a DIFFERENT
    /// process mutating a DIFFERENT layout that happens to share this directory) is actively mid-commit on
    /// can never be mistaken for an orphan and swept out from under it.
    /// </para>
    /// <para>
    /// BEST-EFFORT: every failure — a locked file (another process's antivirus scan, or the narrow window
    /// where a genuinely live commit is mid-<see cref="File.Move(string, string, bool)"/>), an inaccessible
    /// directory, a candidate deleted by a concurrently-running sweep — is swallowed. A missed sweep simply
    /// leaves the file for a later mutating call to retry; it must never fail the edit this cleanup
    /// piggybacks on.
    /// </para>
    /// </remarks>
    private static void SweepStaleStagingFiles(string layoutPath)
    {
        try
        {
            var targetDir = Path.GetDirectoryName(Path.GetFullPath(layoutPath));
            if (string.IsNullOrEmpty(targetDir) || !Directory.Exists(targetDir))
            {
                return;
            }

            var cutoffUtc = DateTime.UtcNow - StageFileRetentionWindow;
            foreach (var glob in StagingFileGlobs)
            {
                foreach (var candidate in Directory.EnumerateFiles(targetDir, glob))
                {
                    try
                    {
                        // CreationTimeUtc, not LastWriteTimeUtc — see StageFileRetentionWindow's remarks.
                        if (File.GetCreationTimeUtc(candidate) < cutoffUtc)
                        {
                            File.Delete(candidate);
                        }
                    }
                    catch
                    {
                        // Best-effort: leave this one candidate for a later sweep rather than failing the edit.
                    }
                }
            }
        }
        catch
        {
            // Best-effort: an inaccessible/unreadable directory must never fail the edit.
        }
    }

    /// <summary>
    /// Returns the members of <paramref name="after"/> that have no matching entry in
    /// <paramref name="baseline"/> (matching one-for-one on equal strings, so a duplicate in
    /// <paramref name="after"/> only "cancels" against a distinct baseline entry) — i.e. what this edit
    /// newly introduced. Mirrors the identical BEFORE-vs-AFTER reconciliation the OpenXmlValidator gate in
    /// <see cref="GuardMutate{TResult}"/> does for structural-error descriptions.
    /// </summary>
    private static List<string> DiffOut(IEnumerable<string> baseline, IEnumerable<string> after)
    {
        var remaining = new List<string>(baseline);
        var added = new List<string>();
        foreach (var item in after)
        {
            var idx = remaining.IndexOf(item);
            if (idx >= 0)
            {
                remaining.RemoveAt(idx);
            }
            else
            {
                added.Add(item);
            }
        }

        return added;
    }

    /// <summary>
    /// Table-grid-violation-specific counterpart of <see cref="DiffOut"/> (see
    /// <see cref="TableGridViolation"/>'s own remarks for the full rationale). Plain description-string
    /// diffing misfires here because <see cref="TableGridViolation.Reason"/> embeds the LIVE grid/coverage
    /// counts, which <c>insert_column</c>/<c>remove_column</c> legitimately change for every row of a table —
    /// so a pre-existing violation's description changes even though the same physical damage persists,
    /// reading as newly introduced.
    /// </summary>
    /// <remarks>
    /// Groups both sets by <c>(</c><see cref="TableGridViolation.Part"/><c>,</c>
    /// <see cref="TableGridViolation.TableIndex"/><c>,</c> <see cref="TableGridViolation.Kind"/><c>)</c> —
    /// deliberately NOT <see cref="TableGridViolation.RowIndex"/> — and flags a key only when the AFTER count
    /// for it EXCEEDS the BEFORE count, returning that many of the AFTER group's violations (the last N in
    /// document order — an arbitrary but stable choice, since which specific row is "the new one" is not
    /// knowable from counts alone) as the representative offenders for the error message. Dropping RowIndex
    /// from the key is what tolerates a row-level <c>remove_control</c> shifting a later, untouched violation's
    /// row index (see <see cref="TableGridViolation"/>'s remarks for the full reasoning, including the
    /// accepted TableIndex-shift imprecision this same choice does NOT extend to).
    /// </remarks>
    private static List<string> DiffTableGridViolations(
        IReadOnlyList<TableGridViolation> baseline, IReadOnlyList<TableGridViolation> after)
    {
        var baselineCounts = new Dictionary<(string Part, int TableIndex, TableGridViolationKind Kind), int>();
        foreach (var v in baseline)
        {
            var key = (v.Part, v.TableIndex, v.Kind);
            baselineCounts[key] = baselineCounts.GetValueOrDefault(key) + 1;
        }

        var added = new List<string>();
        foreach (var group in after.GroupBy(v => (v.Part, v.TableIndex, v.Kind)))
        {
            var afterList = group.ToList();
            var newCount = afterList.Count - baselineCounts.GetValueOrDefault(group.Key);
            if (newCount > 0)
            {
                added.AddRange(afterList.Skip(afterList.Count - newCount).Select(v => v.Describe()));
            }
        }

        return added;
    }

    /// <summary>Thin <see cref="GuardMutate{TResult}"/> instantiation building the content-control edit tools' <see cref="EditResultDto"/> — used by <see cref="EditTools"/>'s <c>insert_field</c>/<c>insert_label</c>/<c>insert_picture</c>/<c>remove_control</c> and <see cref="TableTools"/>'s <c>insert_repeater_table</c>.</summary>
    internal static ToolResponse GuardEdit(string layoutPath, Func<WordprocessingDocument, EditResult> mutate) =>
        GuardMutate(layoutPath, mutate, (editResult, quick) =>
        {
            var dto = new EditResultDto(
                editResult.Operation,
                editResult.ControlId,
                editResult.Alias,
                editResult.XPath,
                editResult.Kind,
                editResult.ColumnCount,
                editResult.Part,
                editResult.Summary,
                new ValidationSummaryDto("quick", quick.Passed, quick.ErrorCount, quick.WarningCount),
                editResult.TableIndex,
                editResult.DataRowIndex);

            return ToolResponse.Success(dto);
        });

    /// <summary>
    /// Thin <see cref="GuardMutate{TResult}"/> instantiation building <see cref="EditTools"/>'s two plain-text
    /// cell edit tools' (<c>set_cell_text</c>, <c>clear_cell_text</c>) <see cref="CellEditResultDto"/> — the
    /// same save-or-reject safety and the same structural QuickValidation guarantee as <see cref="GuardEdit"/>,
    /// for a result shape addressed by (table, row, column) rather than a control id.
    /// </summary>
    internal static ToolResponse GuardCellEdit(string layoutPath, Func<WordprocessingDocument, CellEditResult> mutate) =>
        GuardMutate(layoutPath, mutate, (cellResult, quick) =>
        {
            var dto = new CellEditResultDto(
                cellResult.Operation,
                cellResult.Part,
                cellResult.TableIndex,
                cellResult.Row,
                cellResult.Col,
                cellResult.PreviousText,
                cellResult.NewText,
                cellResult.Summary,
                new ValidationSummaryDto("quick", quick.Passed, quick.ErrorCount, quick.WarningCount));

            return ToolResponse.Success(dto);
        });

    /// <summary>
    /// Thin <see cref="GuardMutate{TResult}"/> instantiation for <see cref="TableTools"/>'s table-STRUCTURE
    /// tools (<c>set_column_widths</c>, <c>insert_column</c>, <c>remove_column</c>, <c>merge_cells</c>,
    /// <c>split_cells</c>), building <see cref="TableEditResultDto"/>. Same save-or-reject safety and
    /// post-edit QuickValidation guarantee as <see cref="GuardEdit"/>/<see cref="GuardCellEdit"/>, plus the
    /// <see cref="TableGridConsistencyGuard"/> backstop <see cref="GuardMutate{TResult}"/> now runs for every
    /// mutating tool — so a structural table edit that would desync a row's cells from the grid is rejected
    /// with the file untouched.
    /// </summary>
    internal static ToolResponse GuardTableEdit(string layoutPath, Func<WordprocessingDocument, TableEditResult> mutate) =>
        GuardMutate(layoutPath, mutate, (tableResult, quick) =>
        {
            var dto = new TableEditResultDto(
                tableResult.Operation,
                tableResult.Part,
                tableResult.TableIndex,
                tableResult.ColumnIndex,
                tableResult.RowsAffected,
                tableResult.ColumnCountBefore,
                tableResult.ColumnCountAfter,
                tableResult.Summary,
                new ValidationSummaryDto("quick", quick.Passed, quick.ErrorCount, quick.WarningCount));

            return ToolResponse.Success(dto);
        });

    /// <summary>
    /// Thin <see cref="GuardMutate{TResult}"/> instantiation for <see cref="LifecycleTools"/>'s
    /// <c>refresh_xml_part</c>, building <see cref="RefreshResultDto"/>. Reuses
    /// <see cref="RefreshResult.QuickValidation"/> (computed by <see cref="LayoutRefresher.Refresh"/> itself,
    /// right after its own mutations) rather than the <see cref="ValidationResult"/>
    /// <see cref="GuardMutate{TResult}"/> passes alongside it — both reflect the identical post-mutate
    /// document state, so either is correct; this one is simply already at hand.
    /// </summary>
    internal static ToolResponse GuardRefresh(string layoutPath, Func<WordprocessingDocument, RefreshResult> mutate) =>
        GuardMutate(layoutPath, mutate, (result, _) =>
        {
            var dto = new RefreshResultDto(
                result.OldReportName,
                result.OldReportId,
                result.OldNamespace,
                result.NewReportName,
                result.NewReportId,
                result.NewNamespace,
                result.StoreItemId,
                result.NamespaceChanged,
                result.RemappedCount,
                result.OrphanedBindings.Select(o => new OrphanedBindingDto(o.Alias, o.XPath, o.Part, o.SdtId)).ToList(),
                result.NewUnboundFields,
                new ValidationSummaryDto(
                    result.QuickValidation.Level,
                    result.QuickValidation.Passed,
                    result.QuickValidation.ErrorCount,
                    result.QuickValidation.WarningCount));

            return ToolResponse.Success(dto);
        });

    // ---- read-tool edit-lock coordination ----

    /// <summary>
    /// Runs a READ-ONLY tool body (<see cref="ReadTools"/>'s <c>get_layout_info</c>/<c>list_dataset_fields</c>/
    /// <c>validate_layout</c>) under the SAME per-path lock pair every mutating tool takes via
    /// <see cref="GuardMutate{TResult}"/> — the in-process lock (<see cref="EditLockFor"/>) THEN the cross-
    /// process mutex (<see cref="CrossProcessLock"/>) INSIDE it — so a read can never observe a layout mid-
    /// commit and a concurrent edit's atomic rename can never be blocked by an open read handle.
    /// </summary>
    /// <remarks>
    /// WHY A READ NEEDS THIS AT ALL: <see cref="GuardMutate{TResult}"/>'s commit
    /// is <see cref="File.Move(string, string, bool)"/> replacing <c>layoutPath</c> in place. On Windows,
    /// replacing a file that ANOTHER open handle holds without <see cref="FileShare.Delete"/> throws a
    /// sharing-violation <see cref="IOException"/> — and every read tool opens the layout with
    /// <see cref="DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(string, bool)"/>'s default
    /// share mode, which does not grant <see cref="FileShare.Delete"/>. Before this method existed, an agent
    /// pipelining <c>insert_field</c> immediately followed by <c>get_layout_info</c> against the SAME file
    /// could make the EDIT's own commit fail with a sharing violation caused by the READ that ran concurrently
    /// with it — misreported as a generic <c>internal_error</c> on the edit, even though the file itself was
    /// never touched by that read. Taking the identical lock pair here closes that race the same way two
    /// mutating tools already avoid contending over one working-copy/stage-and-swap sequence.
    /// <para>
    /// SHORT TIMEOUTS, NOT INDEFINITE BLOCKING: unlike <see cref="GuardMutate{TResult}"/>'s in-process
    /// <c>lock</c> statement (which blocks indefinitely — a mutating edit legitimately may need to wait behind
    /// a slow <c>preview_layout</c> call on the same path, up to its Word-COM conversion timeout), a READ backs
    /// off quickly via <see cref="Monitor.TryEnter(object, TimeSpan)"/> and <see cref="CrossProcessLock.ReadTimeout"/>
    /// and reports <c>file_locked</c> instead — a read is not worth blocking for as long as a slow write might
    /// run; the agent can simply retry it moments later. Both timeouts are independently bounded, so the worst
    /// case total wait is bounded by their sum, not unbounded.
    /// </para>
    /// <para>
    /// READ CONCURRENCY FOR DIFFERENT PATHS IS UNCHANGED: both <see cref="EditLockFor"/> (keyed per full path)
    /// and <see cref="CrossProcessLock"/> (mutex name keyed per full path's hash) key strictly per-path, so two
    /// reads (or a read and an edit) targeting DIFFERENT layouts never contend here — only same-path traffic
    /// is serialized, exactly as for the mutating tools.
    /// </para>
    /// </remarks>
    internal static ToolResponse GuardRead(string layoutPath, Func<ToolResponse> body)
    {
        return Guard(() =>
        {
            var monitor = EditLockFor(layoutPath);
            if (!Monitor.TryEnter(monitor, CrossProcessLock.ReadTimeout))
            {
                return CrossProcessLockTimeoutFailure(layoutPath, CrossProcessLock.ReadTimeout);
            }

            try
            {
                using var crossLock = CrossProcessLock.TryAcquire(layoutPath, CrossProcessLock.ReadTimeout);
                if (!crossLock.Acquired)
                {
                    return CrossProcessLockTimeoutFailure(layoutPath, CrossProcessLock.ReadTimeout);
                }

                return body();
            }
            finally
            {
                Monitor.Exit(monitor);
            }
        }, layoutPath);
    }

    /// <summary>
    /// Shared <c>file_locked</c> envelope for a timed-out lock wait — used by <see cref="GuardMutate{TResult}"/>,
    /// <c>LifecycleTools.PreviewLayout</c>/<c>CreateLayout</c>, and <see cref="GuardRead"/> alike, so every
    /// flavor of "this layout is busy right now" (whether the wait was on the in-process lock or the cross-
    /// process mutex) reports the identical error code/hint shape as the sharing-violation <c>IOException</c>
    /// branch in <see cref="Guard"/> below — a caller need only ever branch on ONE code, <c>file_locked</c>,
    /// regardless of which of the three causes actually produced it.
    /// </summary>
    internal static ToolResponse CrossProcessLockTimeoutFailure(string layoutPath, TimeSpan timeout) =>
        ToolResponse.Failure(
            "file_locked",
            $"Timed out after {timeout.TotalSeconds:0.#}s waiting for the edit lock on "
            + $"'{Path.GetFullPath(layoutPath)}'.",
            "Another process (e.g. another IDE window's MCP host) is currently editing or previewing this "
            + "layout - wait for it to finish and retry. If nothing else is actually using this file, a prior "
            + "host process may have exited while holding the lock; it is released automatically, so a retry "
            + "after a short pause should succeed.");

    // ---- location building (shared by EditTools' insert_field/insert_label and TableTools' insert_repeater_table) ----

    /// <summary>
    /// Builds a <see cref="Location"/> from a tool's flat, MCP-friendly parameters. Only checks that
    /// <paramref name="locationType"/> and <paramref name="layoutPart"/> themselves name a known
    /// <see cref="LocationKind"/>/<see cref="LayoutPart"/> (case-insensitively); per-kind required-field
    /// checks (e.g. <c>afterControl</c> needing <paramref name="controlId"/>) are left to
    /// <see cref="Location.Validate"/>, which <see cref="LocationResolver.Resolve"/> always calls before
    /// use — one <see cref="ArgumentException"/> source instead of two. Whether <paramref name="partName"/>
    /// actually names a part present in THIS layout is a "does it resolve against this document" question,
    /// left to <see cref="LocationResolver.Resolve"/> too (surfaced as <c>not_found</c>, not
    /// <c>invalid_argument</c> — see <see cref="NotFoundHintFor"/>).
    /// </summary>
    internal static Location BuildLocation(
        string locationType, int? controlId, int? tableIndex, int? row, int? col, string? searchText,
        string layoutPart = "body", string? partName = null)
    {
        if (!TryParseLocationKind(locationType, out var kind))
        {
            throw new ArgumentException(
                $"Unknown locationType '{locationType}'. Use 'documentEnd', 'afterControl', 'tableCell', or "
                + "'atText'.",
                nameof(locationType));
        }

        if (!TryParseLayoutPart(layoutPart, out var part))
        {
            throw new ArgumentException(
                $"Unknown layoutPart '{layoutPart}'. Use 'body', 'header', or 'footer'.",
                nameof(layoutPart));
        }

        return new Location
        {
            Type = kind,
            Part = part,
            PartName = string.IsNullOrWhiteSpace(partName) ? null : partName,
            ControlId = controlId,
            TableIndex = tableIndex,
            Row = row,
            Col = col,
            SearchText = searchText,
        };
    }

    private static bool TryParseLocationKind(string locationType, out LocationKind kind)
    {
        switch (locationType.Trim().ToLowerInvariant())
        {
            case "documentend":
                kind = LocationKind.DocumentEnd;
                return true;
            case "aftercontrol":
                kind = LocationKind.AfterControl;
                return true;
            case "tablecell":
                kind = LocationKind.TableCell;
                return true;
            case "attext":
                kind = LocationKind.AtText;
                return true;
            default:
                kind = default;
                return false;
        }
    }

    /// <summary>
    /// Parses a <see cref="LayoutPart"/> from a tool's flat string parameter (case-insensitive). Internal
    /// rather than private: shared by <see cref="BuildLocation"/> here and by <see cref="TableTools"/>'s own
    /// <c>ParseLayoutPartOrThrow</c> (the table-structure tools address a table directly by
    /// <see cref="LayoutPart"/> rather than building a full <see cref="Location"/>).
    /// </summary>
    internal static bool TryParseLayoutPart(string layoutPart, out LayoutPart part)
    {
        switch (layoutPart.Trim().ToLowerInvariant())
        {
            case "body":
                part = LayoutPart.Body;
                return true;
            case "header":
                part = LayoutPart.Header;
                return true;
            case "footer":
                part = LayoutPart.Footer;
                return true;
            default:
                part = default;
                return false;
        }
    }

    // ---- mapping helpers (shared response-DTO construction) ----

    internal static HashSet<string> BuildBoundPaths(LayoutInventory inventory)
    {
        var set = new HashSet<string>(StringComparer.Ordinal);
        foreach (var control in inventory.Controls)
        {
            if (control.XPath is null)
            {
                continue;
            }

            var segments = BindingXPath.Segments(control.XPath);
            if (segments.Count <= 1)
            {
                continue;
            }

            // Drop the root segment; the remainder is the column/item path e.g. /Header/Line/ItemNo_Line.
            set.Add("/" + string.Join("/", segments.Skip(1)));
        }

        return set;
    }

    internal static ReportInfoDto ToReportDto(ReportIdentity r) =>
        new(r.ReportName, r.ReportId, r.Namespace, r.StoreItemId);

    internal static ControlDto ToControlDto(LayoutControl c) =>
        new(c.Kind.ToString(), c.Alias, c.Tag, c.XPath, c.StoreItemId, c.Part, c.SdtId, c.UsesW15Binding,
            c.ParentRepeater?.XPath, ToLevelString(c.Level), c.TableIndex, c.RowIndex, c.ColIndex);

    /// <summary>Camel-cases the <see cref="SdtLevel"/> enum name for the JSON contract (run/block/cell/row/runRuby/unknown).</summary>
    private static string ToLevelString(SdtLevel level) => level switch
    {
        SdtLevel.Run => "run",
        SdtLevel.Block => "block",
        SdtLevel.Cell => "cell",
        SdtLevel.Row => "row",
        SdtLevel.RunRuby => "runRuby",
        _ => "unknown",
    };

    internal static PartInfoDto ToPartInfoDto(LayoutPartInfo p) =>
        new(p.Name, ToPartKindString(p.Kind), ToRoleString(p.Role), p.IsDefaultTarget);

    /// <summary>Camel-cases the <see cref="LayoutPartKind"/> enum name for the JSON contract (document/header/footer).</summary>
    private static string ToPartKindString(LayoutPartKind kind) => kind switch
    {
        LayoutPartKind.Header => "header",
        LayoutPartKind.Footer => "footer",
        _ => "document",
    };

    /// <summary>Camel-cases the <see cref="HeaderFooterRole"/> enum name for the JSON contract (default/first/even, null when unreferenced).</summary>
    private static string? ToRoleString(HeaderFooterRole? role) => role switch
    {
        HeaderFooterRole.Default => "default",
        HeaderFooterRole.First => "first",
        HeaderFooterRole.Even => "even",
        _ => null,
    };

    internal static TableDto ToTableDto(TableStructure t) =>
        new(t.Part, t.TableIndex, t.RowCount, t.ColumnCount, t.GridColumnWidths,
            t.Rows.Select(r => new TableRowDto(
                r.RowIndex,
                r.IsControlRow,
                r.ControlId,
                r.Cells.Select(cell => new TableCellDto(
                    cell.ColIndex,
                    cell.IsControlCell,
                    cell.ControlId,
                    cell.ControlKind,
                    cell.Alias,
                    cell.XPath,
                    cell.Text,
                    cell.InnerControlIds)).ToList())).ToList());

    /// <summary>
    /// SECURITY: recurses once per <see cref="DataItem"/>
    /// nesting level with no depth counter of its own — deliberately, for the same reason
    /// <c>BcWordLayout.Merge.SampleDataGenerator.BuildInstance</c> needs none: every <see cref="DataItem"/>
    /// tree is built by <c>BcWordLayout.Domain.SchemaProvider.BuildNode</c>, which already rejects (via
    /// <c>BcWordLayout.Domain.ResourceLimits.MaxSchemaDepth</c>) any tree deeper than the cap before this
    /// method ever sees it, so a second counter here would be redundant.
    /// </summary>
    internal static DataItemDto ToDataItemDto(DataItem item, HashSet<string>? boundPaths)
    {
        var columns = item.Columns
            .Select(col => new ColumnDto(
                col.Name,
                col.Path,
                col.IsLabel,
                boundPaths is null ? null : boundPaths.Contains(col.Path)))
            .ToList();

        var children = item.Children.Select(child => ToDataItemDto(child, boundPaths)).ToList();
        return new DataItemDto(item.Name, item.Path, item.IsSystem, columns, children);
    }

    // ---- exception-to-envelope translation ----

    /// <summary>
    /// Runs a tool body, translating known exceptions into the structured error envelope so nothing throws
    /// across the MCP boundary. Every branch supplies a specific, non-empty <see cref="ToolError.Hint"/> -
    /// required by <see cref="ToolResponse.Failure"/>'s signature (: see its own remarks) - and, for
    /// the codes with more than one realistic cause (<c>invalid_argument</c>, <c>not_found</c>), tailors that
    /// hint to the SPECIFIC argument or lookup target that failed (see
    /// <see cref="InvalidArgumentHint"/>/<see cref="NotFoundHintFor"/>) instead of repeating one generic
    /// sentence regardless of cause.
    /// </summary>
    /// <remarks>
    /// <c>not_found</c> is produced ONLY by a caught <see cref="NotFoundException"/> - a dedicated Domain
    /// type constructed deliberately at genuine lookup-failure throw sites (see
    /// <see cref="NotFoundException"/>'s own remarks for why that distinction had to stop resting
    /// on <see cref="InvalidOperationException"/>, the BCL's default failure type). A raw
    /// <see cref="InvalidOperationException"/> - e.g. an internal "this should be impossible" assertion, or a
    /// stray LINQ <c>.First()</c>/<c>.Single()</c> - is NOT specially handled here; it falls through to the
    /// generic <c>catch (Exception)</c> below and reports <c>internal_error</c>, which is what an unreviewed
    /// assertion should report rather than misleading the calling agent into a futile "retry with different
    /// arguments" loop.
    /// <para>
    /// Similarly, <c>invalid_layout</c>'s hint is tailored by catching <see cref="ResourceLimitExceededException"/>
    /// EXPLICITLY, before the generic <see cref="InvalidDataException"/> branch (
    /// addendum: see <see cref="ResourceLimitExceededException"/>'s own remarks) - a size/part-count/nesting-
    /// depth rejection gets its own "this file is too big/deeply nested" hint instead of the generic "missing
    /// dataset part/wrong namespace" one, again keyed off the TYPE rather than any <c>ex.Message</c> text.
    /// </para>
    /// <para>
    /// <c>file_locked</c> is produced two ways: (1) below, by catching
    /// <see cref="IOException"/> and checking <see cref="Exception.HResult"/> for a Windows sharing/lock
    /// violation (<c>ERROR_SHARING_VIOLATION</c> 0x80070020 / <c>ERROR_LOCK_VIOLATION</c> 0x80070021) - the
    /// "this layout is open in Word (or another program) right now" case, where opening/copying the file
    /// itself fails; and (2) not through this method at all, but as a plain <see cref="ToolResponse.Failure"/>
    /// returned directly by <see cref="CrossProcessLockTimeoutFailure"/> when the cross-process/in-process
    /// edit-lock WAIT itself times out (see <see cref="GuardMutate{TResult}"/>, <see cref="GuardRead"/>, and
    /// <c>LifecycleTools.PreviewLayout</c>/<c>CreateLayout</c>) - the "another process/thread is currently
    /// editing this same layout" case, which is not an exception at all and so never reaches this catch chain.
    /// Both land on the identical <c>file_locked</c> code so a caller need only branch on one value regardless
    /// of which of the two actually happened. An <see cref="IOException"/> whose <see cref="Exception.HResult"/>
    /// does NOT match either code (e.g. a genuinely corrupt/unreadable file, or - on a non-Windows OS - any
    /// I/O failure at all, since these specific HRESULT values are a Windows convention) falls through to the
    /// generic <c>catch (Exception)</c> below exactly as before this change, keeping today's <c>internal_error</c>
    /// behavior for every OTHER I/O failure shape.
    /// </para>
    /// <para>
    /// <c>template_not_unbound</c> is produced ONLY by a caught
    /// <see cref="TemplateNotUnboundException"/> - <c>create_layout</c>'s <c>templatePath</c> already carried
    /// its own BC dataset part AND bound content controls that go stale the moment that part is replaced (see
    /// that type's own remarks for why this refuses rather than reports a warning on an otherwise successful
    /// result, as it originally did). Caught explicitly, ahead of the generic
    /// <see cref="InvalidDataException"/>/<c>catch (Exception)</c> branches, exactly like
    /// <see cref="ResourceLimitExceededException"/> above - keyed off the TYPE, never <c>ex.Message</c> text.
    /// </para>
    /// </remarks>
    internal static ToolResponse Guard(Func<ToolResponse> body, string path)
    {
        try
        {
            return body();
        }
        catch (FileNotFoundException ex)
        {
            // Every throw site in this codebase gives ex.Message its own specific wording naming WHICH
            // parameter was checked (e.g. "layoutPath does not point to an existing file.", "schemaSource
            // does not point to an existing file." - see e.g. LayoutBuilder.Create's own schemaSource vs
            // templatePath split). The actual path value is appended here, centrally, from ex.FileName
            // (every throw site passes it explicitly) falling back to the path this method itself was called
            // with - which still disambiguates correctly for a tool with more than one candidate path (e.g.
            // create_layout's schemaSource vs templatePath), where the single `path` passed to this method
            // may not be the one that was actually missing.
            var missingPath = ex.FileName ?? path;
            return ToolResponse.Failure(
                "file_not_found",
                $"{ex.Message} Path checked: '{missingPath}'.",
                "Pass an absolute path (not relative) to a file that already exists on disk - the parameter "
                + "named above must point at it; check for typos, a missing drive letter, or a relative path "
                + "that does not resolve the way you expect.");
        }
        catch (IOException ex) when (IsSharingOrLockViolation(ex))
        {
            // "This layout is open in Word (or another program) right now" - the
            // sharing-violation half of the file_locked story (see this method's own remarks above for the
            // other half - a timed-out edit-lock WAIT, which never reaches this catch chain at all).
            return ToolResponse.Failure(
                "file_locked",
                ex.Message,
                "This file is open in another program (typically Word) or is being read/written by another "
                + "process right now - close it there (or wait for that other operation to finish) and retry; "
                + "the layout on disk is unaffected by this failure.");
        }
        catch (ResourceLimitExceededException ex)
        {
            // ResourceLimitExceededException does NOT derive from InvalidDataException (that BCL type is
            // sealed - see the exception's own remarks), so this branch and the plain InvalidDataException
            // one below never actually compete for the same thrown instance; it is listed first purely for
            // readability (most-specific-first). A size/part-count/nesting-depth rejection (
            // gets its OWN tailored hint here, keyed off the TYPE (never by pattern-
            // matching ex.Message substrings - the exact string-coupling already eliminated
            // for not_found/invalid_argument). Without this branch such a rejection would fall through to the
            // generic catch (Exception) below and misreport as internal_error instead of invalid_layout.
            return ToolResponse.Failure(
                "invalid_layout",
                ex.Message,
                "This file exceeds one of this tool's supported size/nesting limits (a custom XML part or "
                + "schema file too large, too many custom XML parts, or schema/document nesting too deep) - "
                + "it is likely malformed or a maliciously crafted file rather than a real Business Central "
                + "layout/schema. This is not something to retry with the same file.");
        }
        catch (InvalidDataException ex)
        {
            return ToolResponse.Failure(
                "invalid_layout",
                ex.Message,
                "Confirm the file is either a BC Word layout (.docx) containing a dataset custom XML part, "
                + "or a standalone exported schema .xml whose root is 'NavWordReportXmlPart' in the "
                + "'urn:microsoft-dynamics-nav/reports/<ReportName>/<id>/' namespace; use get_layout_info "
                + "(for a .docx) to inspect what the file actually contains.");
        }
        catch (TemplateNotUnboundException ex)
        {
            // create_layout's templatePath already carried its own BC dataset part
            // AND bound content controls that would go stale the moment that part is replaced - refused
            // outright by LayoutBuilder.Create rather than reported as a warning on an Ok=true result (see
            // that exception type's own remarks). Nothing was written - see LayoutBuilder.Create's own
            // atomic-write behavior.
            return ToolResponse.Failure(
                "template_not_unbound",
                ex.Message,
                "templatePath must be an UNBOUND branded/styled shell (headers/footers, logo, fonts/styles "
                + "only) - not a full BC layout with its own bound content controls. If you want to keep "
                + "THIS layout's design but bind it to a new schema, copy the layout to a new path and call "
                + "refresh_xml_part on the COPY instead (it keeps the layout's own storeItemID and remaps "
                + "existing bindings by element name, rather than replacing the part outright). If you "
                + "specifically want to reuse it as a branded shell for create_layout, first strip its bound "
                + "controls with remove_control (repeat for every control get_layout_info lists) or supply a "
                + "genuinely unbound template, then retry.");
        }
        catch (ArgumentException ex)
        {
            return ToolResponse.Failure("invalid_argument", ex.Message, InvalidArgumentHint(ex.ParamName));
        }
        catch (NotFoundException ex)
        {
            return ToolResponse.Failure("not_found", ex.Message, NotFoundHintFor(ex.TargetKind));
        }
        catch (Exception ex)
        {
            return ToolResponse.Failure(
                "internal_error",
                ex.Message,
                "This is an unexpected failure, not one of this tool's normal validation outcomes. Confirm "
                + "the file is a valid, uncorrupted OOXML .docx (e.g. it opens in Word); if the same failure "
                + "persists against a known-good file, this likely indicates a bug in the tool itself and is "
                + "worth reporting together with the full error message above.");
        }
    }

    /// <summary>
    /// Recognizes the two Windows sharing/lock-violation <see cref="Exception.HResult"/> values a caller-
    /// blocking <see cref="IOException"/> carries when the underlying <c>CreateFile</c>/rename call failed
    /// because ANOTHER open handle (typically Word, holding the file the user has it open in) did not grant
    /// the share mode this operation needed - <c>ERROR_SHARING_VIOLATION</c> (0x20) when opening/copying the
    /// file itself, <c>ERROR_LOCK_VIOLATION</c> (0x21) when a byte-range lock conflicts (rarer for whole-file
    /// OOXML access, checked for completeness). Both are surfaced as an <see cref="IOException"/> whose
    /// <see cref="Exception.HResult"/> is <c>HRESULT_FROM_WIN32(code)</c> = <c>0x80070000 | code</c> - the
    /// standard Win32-error-wrapped-as-HRESULT convention every BCL I/O API uses on Windows. On a non-Windows
    /// OS these specific HRESULT values simply never occur (a locked-file error there has a different HRESULT
    /// entirely, if it is even representable as one), so this predicate naturally never matches there and
    /// Guard's ordinary <c>catch (Exception)</c> -&gt; <c>internal_error</c> path is unaffected - no runtime
    /// OS check is needed to keep this Windows-specific without breaking other platforms.
    /// </summary>
    private static bool IsSharingOrLockViolation(IOException ex)
    {
        const int ErrorSharingViolation = unchecked((int)0x80070020);
        const int ErrorLockViolation = unchecked((int)0x80070021);
        return ex.HResult is ErrorSharingViolation or ErrorLockViolation;
    }

    /// <summary>
    /// Builds an <c>invalid_argument</c> hint tailored to the specific argument that failed, keyed off
    /// <see cref="ArgumentException.ParamName"/> (populated via <c>nameof(...)</c> at every throw site across
    /// this tool surface - <c>ToolGuards</c>/the tool-family classes themselves, <c>SdtFactory</c>,
    /// <c>LayoutEditor</c>, <c>Location.Validate</c>, <c>LocationResolver</c>, <c>TableStructureEditor</c>,
    /// <c>CellTextEditor</c>, <c>LayoutBuilder</c>, and <c>LayoutRefresher</c>): names that argument's valid
    /// values/shape and, where useful, the inspection tool that reports them, rather than one generic
    /// sentence for every <c>invalid_argument</c> failure regardless of which argument actually caused it.
    /// Falls back to a hint covering every possibility when <paramref name="paramName"/> is null (not every
    /// <see cref="ArgumentException"/> a caller could throw sets it - <c>TableStructureEditor</c>'s
    /// vMerge rejection is the one live example) or is not one this map recognizes.
    /// <para>
    /// GUARANTEE: every branch below is pinned by a named test in
    /// <c>InvalidArgumentHintCoverageTests</c>, which drives each key end-to-end through the real tool
    /// surface (or, for the handful of branches no documented tool call can reach - a defensively-checked
    /// invariant, or a scenario only reachable via a hand-crafted document - through the same public API the
    /// throw site itself lives on, noted at that test). Renaming or retiring a <c>nameof(...)</c> at any of
    /// the files above therefore fails a SPECIFIC named test instead of silently rerouting that failure to
    /// the generic fallback below with no compile/test signal - the same guarantee <see cref="NotFoundHintFor"/>
    /// already has for <c>not_found</c> (see its own remarks), now extended to <c>invalid_argument</c>. Two
    /// classes of drift this once let through, found by that coverage sweep and fixed here: a DEAD key (a
    /// case label no throw site anywhere ever produces) - <c>"row"</c>/<c>"col"</c> stood in for
    /// <c>"tableIndex"</c>, which every real throw site actually reports even when Row/Col specifically is
    /// the invalid field, and <c>"field"</c>/<c>"label"</c>/<c>"dataItem"</c> stood in for those tool
    /// parameters' own names, which are never the reported <c>ParamName</c> - validation for all three
    /// happens one layer down in <c>SdtFactory</c> under <c>"datasetPath"</c>/<c>"dataItemPath"</c> instead;
    /// both dead sub-keys were removed. And an UNMAPPED real <c>ParamName</c> (a live, tool-reachable throw
    /// site whose name matched no case at all, silently degrading to the generic fallback) - <c>"schemaSource"</c>/
    /// <c>"outputPath"</c> (<c>create_layout</c>) and <c>"newSchemaSource"</c> (<c>refresh_xml_part</c>) were
    /// found unmapped and given their own case below.
    /// </para>
    /// </summary>
    private static string InvalidArgumentHint(string? paramName) => paramName?.ToLowerInvariant() switch
    {
        "locationtype" =>
            "locationType must be one of: 'documentEnd', 'afterControl', 'tableCell', or 'atText' "
            + "(case-insensitive).",

        "layoutpart" =>
            "layoutPart must be one of: 'body' (default), 'header', or 'footer' (case-insensitive); "
            + "partName (optional) only applies when layoutPart is 'header'/'footer' and names a specific "
            + "part file (e.g. 'header2.xml') — omit it to target the DEFAULT header/footer part (see "
            + "get_layout_info's partDetails for each part's role).",

        "location" =>
            "The location you targeted is not valid for this operation (see the message above for which "
            + "applies). For set_cell_text/clear_cell_text: the cell must be PLAIN TEXT - a cell that holds a "
            + "bound field/label control is rejected; use remove_control (the cell/column is preserved) or "
            + "insert_field/insert_label for those. For insert_repeater_table: only layoutPart='body' (the "
            + "default) is supported - a repeater TABLE in a header/footer is deferred (GitHub issue #10) "
            + "(unlike insert_field/insert_label, which do support 'header'/'footer').",

        "controlid" =>
            "afterControl requires controlId: the w:id of an existing control in this layout - see "
            + "get_layout_info's control inventory (controls[].sdtId) for the real ids present.",

        "tableindex" =>
            "tableCell requires non-negative tableIndex, row, and col (all 0-based); call get_layout_info, "
            + "or open the file, to see how many tables/rows/cells actually exist before retrying.",

        "searchtext" =>
            "atText requires a non-empty searchText substring that actually appears in a run of text in the "
            + "document body; inspect the layout's visible text (e.g. via get_layout_info or by opening the "
            + "file) and retry with an exact substring that is really present.",

        "datasetpath" or "dataitempath" =>
            "The dataset path must name a real column/data item exactly as reported by list_dataset_fields, "
            + "e.g. '/Header/Line/ItemNo_Line'. insert_field/insert_label each require a LEAF column (not a "
            + "repeating data item): insert_label additionally requires a label-shaped name per the active "
            + $"label convention ({LabelConvention.Current.Describe()}), insert_field requires the reverse. "
            + "insert_repeater_table's dataItem must be a repeating, non-system DATA ITEM (not a leaf column).",

        "columns" or "column" =>
            "insert_repeater_table's columns must be a non-empty, comma-separated list of LEAF column names "
            + "belonging to the given dataItem (not full paths, not nested data items) - see "
            + "list_dataset_fields for the exact column names available under that data item.",

        "columnwidths" or "options" =>
            "columnWidths, if supplied, must be a comma-separated list of integers with exactly one width "
            + "(in twips) per column, matching the column count exactly - omit it entirely for an even "
            + "default width per column instead.",

        "widths" =>
            "set_column_widths' widths must be a non-empty comma-separated list of integers (twips), exactly "
            + "one per GRID column - the count must equal the table's columnCount from get_layout_info.",

        "atcolumn" =>
            "insert_column's atColumn is the 0-based GRID position for the new column (0..columnCount from "
            + "get_layout_info; columnCount, or omitting it, appends at the far-right edge). An INTERIOR "
            + "position must land on a cell boundary in every row that carries the new column's content (a "
            + "header row, and the bound data row) - if it falls inside a spanned cell there, split that "
            + "cell first with split_cells or pick a different position.",

        "gridcolumn" =>
            "remove_column's column is a 0-based GRID column index (0..columnCount-1 from get_layout_info); "
            + "the last remaining column cannot be removed.",

        "fromcolumn" or "tocolumn" =>
            "merge_cells needs 0 <= fromColumn < toColumn < the row's cell count (see get_layout_info's "
            + "tables[].rows[].cells), both 0-based physical cell indices within the same row; an absorbed "
            + "cell that holds a bound control is rejected (remove_control it first).",

        "cellindex" =>
            "split_cells' cellIndex is the 0-based physical cell index (within the row) of a SPANNED cell "
            + "(gridSpan > 1) to split; a non-spanned cell has nothing to split.",

        "widthmm" or "heightmm" =>
            "insert_picture's widthMm/heightMm are the picture FRAME size in millimetres, each between 1 "
            + "and 500; omit both for the corpus default (30x30 mm, the size real BC company logos use).",

        "embedrelationshipid" =>
            "A picture control's image reference must point at a real image part in the same OOXML part; "
            + "this is set internally by insert_picture, so reaching this message indicates a tool bug "
            + "rather than a mistake in your call.",

        "look" =>
            "insert_repeater_table's look must be 'bc' (the default - no drawn table grid, just a rule under "
            + "the header row, matching every real BC lines table) or 'grid' (an explicit single-line border "
            + "on every edge and between every cell).",

        "edges" =>
            "set_cell_borders' edges must be a comma-separated list of 'top', 'bottom', 'left', 'right' (or "
            + "'all') naming at least one edge - e.g. edges='top' for the rule above a totals row. Edges you "
            + "do not name keep whatever they already had.",

        "style" or "size" =>
            "set_cell_borders' style must be 'single' (draw a line) or 'none' (explicitly clear the named "
            + "edges), and size is the rule thickness in EIGHTHS of a point between 2 and 96 - the BC "
            + "standard is 4 (½ pt), which is also the default.",

        "mode" =>
            "insert_column's mode must be 'field', 'label', or 'plainText'. 'field'/'label' require a dataPath "
            + "(a full dataset path exactly as list_dataset_fields reports, e.g. '/Header/Line/Discount_Line').",

        "datapath" =>
            "insert_column's dataPath must be a full dataset path exactly as list_dataset_fields reports (e.g. "
            + "'/Header/Line/Discount_Line'); mode='label' additionally requires a label-shaped leaf name per "
            + $"the active label convention ({LabelConvention.Current.Describe()}), mode='field' requires the "
            + "reverse.",

        "keeptext" =>
            "remove_control's keepText=true cannot target a repeating-section control (it would orphan its "
            + "row template from any enclosing repeater); call remove_control again with keepText=false (the "
            + "default) to remove the whole repeater instead.",

        "schema" =>
            "Build the schema via SchemaProvider.FromLayout against a .docx layout (not FromSchemaXml "
            + "against a standalone schema .xml) so a real storeItemID is available to bind against; see the "
            + "message above for exactly what was rejected.",

        "schemasource" =>
            "create_layout's schemaSource must be a non-empty absolute path to an existing .docx layout "
            + "(its BC dataset part is reused) or a standalone exported schema .xml; see the message above "
            + "for exactly what was rejected about it.",

        "newschemasource" =>
            "refresh_xml_part's newSchemaSource must be a non-empty absolute path to an existing .docx "
            + "layout (its BC dataset part is reused) or a standalone exported schema .xml - the same shape "
            + "create_layout's schemaSource accepts; see the message above for exactly what was rejected.",

        "outputpath" =>
            "The output path must be a non-empty absolute path (create_layout's outputPath is where the new "
            + "layout .docx is written); see the message above for exactly what was rejected about it.",

        _ =>
            "Check the argument named in the message above against this tool's own parameter descriptions: "
            + "a field/label/dataItem must be a dataset path exactly as reported by list_dataset_fields; "
            + "locationType must be 'documentEnd', 'afterControl', 'tableCell', or 'atText', each with its "
            + "own required fields (controlId for afterControl; tableIndex+row+col for tableCell; "
            + "searchText for atText).",
    };

    /// <summary>
    /// Builds a <c>not_found</c> hint tailored to what kind of lookup target failed, keyed off the typed
    /// <see cref="NotFoundException.TargetKind"/> the throw site set explicitly - unlike the pre-B11 design,
    /// which sniffed <see cref="Exception.Message"/> substrings because <see cref="InvalidOperationException"/>
    /// has no structured equivalent of <see cref="ArgumentException.ParamName"/> to key off (see
    /// <see cref="NotFoundException"/>'s own remarks: rewording a Domain message could silently degrade that
    /// old scheme to the generic fallback with no compile/test signal - switching on the enum instead keys
    /// every branch here off a value the throw site sets explicitly, immune to message rewording). Falls back
    /// to a hint covering every possibility for <see cref="NotFoundTarget.General"/> (not thrown by any site
    /// today - see its own remarks) or any future enum value this switch has not yet been taught about (the
    /// <c>_</c> arm means a NEW member gets the generic hint, not a compile error - update this switch when
    /// adding one).
    /// </summary>
    private static string NotFoundHintFor(NotFoundTarget targetKind) => targetKind switch
    {
        NotFoundTarget.Control =>
            "No control has that id in the targeted layoutPart/partName (it may still exist elsewhere "
            + "in the layout - AfterControl only searches the part you targeted); call get_layout_info "
            + "and check controls[].sdtId together with controls[].part for the real ids present and "
            + "which part each actually lives in (ids are per-document, not sequential or guessable).",

        NotFoundTarget.NamedHeaderFooterPart =>
            "partName does not match any header/footer part in this layout; call get_layout_info and "
            + "check partDetails for the real header/footer file names present (e.g. 'header1.xml') and "
            + "each part's role, or omit partName to target the DEFAULT part.",

        NotFoundTarget.HeaderFooterParts =>
            "This layout has no header/footer parts at all, so layoutPart='header'/'footer' cannot "
            + "resolve; call get_layout_info and check the parts list, or target layoutPart='body' "
            + "(the default) instead.",

        NotFoundTarget.TableCoordinate =>
            "The table/row/col index is out of range; call get_layout_info, or open the file, to "
            + "count the document's actual tables/rows/cells (all indices are 0-based) before retrying "
            + "tableCell.",

        NotFoundTarget.SearchText =>
            "No run of text in the document body contains that searchText substring; inspect the "
            + "layout's actual visible text (e.g. via get_layout_info or by opening the file) and retry "
            + "with an exact substring that is really present.",

        NotFoundTarget.AfterControlPosition =>
            "That control's position cannot safely host a sibling via afterControl (e.g. a row-level "
            + "repeater control); use tableCell addressing instead, or target a different control id - "
            + "see get_layout_info's control inventory.",

        _ =>
            "The referenced control id, table location, or search text does not exist in this layout; "
            + "call get_layout_info to inspect the current control inventory and document structure before "
            + "retrying.",
    };
}
