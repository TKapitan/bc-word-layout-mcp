using System.Xml;
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
    /// Per-path locks serializing concurrent edits targeting the same layout file (keyed by the file system's
    /// own spelling of the path, <see cref="PathIdentity.Canonical"/>, case-insensitively —
    /// <see cref="StringComparer.OrdinalIgnoreCase"/> is the dictionary's own comparer, so lookups need no
    /// manual case-folding). Keying on the resolved path, not just <see cref="Path.GetFullPath(string)"/>, is
    /// what makes a junction, symlink, or subst or mapped drive spelling of a layout take the SAME lock as its
    /// plain spelling (GitHub issue #92). Two <see cref="GuardEdit"/> calls against
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
        EditLocks.GetOrAdd(PathIdentity.Canonical(layoutPath), static _ => new object());

    /// <summary>
    /// Shared choreography for all 20 mutating tools — every tool that writes through <see cref="GuardEdit"/>,
    /// <see cref="GuardCellEdit"/>, <see cref="GuardTableEdit"/> or <see cref="GuardRefresh"/>: copy
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
    /// different mutating tools return different result shapes — <see cref="EditResult"/> for the nine
    /// content-control and repeater edits, <see cref="CellEditResult"/> and <see cref="TableEditResult"/> for the
    /// cell and table-structure edits, <see cref="RefreshResult"/> for <c>refresh_xml_part</c> (which additionally
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
        string layoutPath, Func<WordprocessingDocument, TResult> mutate, Func<TResult, ValidationResult, ToolResponse> onSuccess,
        string? tool = null)
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

                // The file system's own spelling of layoutPath (see PathIdentity): the folder the commit stages
                // into and the path it replaces. Through a junction or a subst or mapped drive this is the same
                // file either way; through a symlink to the FILE, renaming over layoutPath itself would replace
                // the link with a plain copy and leave the real layout - and every other host editing it under
                // its real name, behind the same lock - without this edit (GitHub issue #92).
                var commitPath = PathIdentity.Canonical(layoutPath);

                // Best-effort orphan cleanup: runs on EVERY mutating call,
                // scoped to layoutPath's OWN directory only — see SweepStaleStagingFiles's own remarks — plus
                // this method's own %TEMP% working copies, at most hourly (see SweepStaleEditWorkingCopiesIfDue).
                SweepStaleStagingFiles(commitPath);
                SweepStaleEditWorkingCopiesIfDue();

                // CopyWritable, not a bare File.Copy: a read-only layout (a TFVC/Perforce file that is not
                // checked out) would otherwise yield a read-only working copy that can be neither opened for
                // write below nor deleted afterwards - see ScratchFiles' own remarks (GitHub issue #94).
                var workingCopy = Path.Combine(Path.GetTempPath(), $"{EditWorkingCopyPrefix}{Guid.NewGuid():N}.docx");
                RetryTransientFileDenial(() => ScratchFiles.CopyWritable(layoutPath, workingCopy));
                try
                {
                    TResult mutateResult;
                    ValidationResult quick;

                    // callerPath: a layoutPath that is not a .docx package at all must be reported by ITS name,
                    // not the private working copy's.
                    using (var doc = LayoutPackage.Open(workingCopy, isEditable: true, autoSave: false, callerPath: layoutPath))
                    {
                        // Baseline BEFORE the edit: see the "does NOT reject" remark above.
                        var baselineErrors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc)
                            .Select(e => e.Description)
                            .ToList();

                        // Second baseline: content controls - and field codes, tables, drawings - nested
                        // inside a plain-text control. Word rejects
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
                                ErrorCodes.EditWouldCorrupt,
                                $"This edit would place {newNesting.Count} content control(s), field code(s), "
                                + "table(s) or drawing(s) inside a plain-text content control, which Word rejects "
                                + "as a corrupt document; the file was NOT modified.",
                                $"Offending nesting: {string.Join(" | ", newNesting.Take(5))}. A plain-text "
                                + "content control (the shape BC uses for most header fields, e.g. an address "
                                + "or document-number cell) can only contain runs of text - never another "
                                + "field/label control, a page-number field, a table or a picture. The location "
                                + "you targeted resolves to a spot INSIDE that control. Insert into a plain cell "
                                + "via 'tableCell' (one that is not itself a control - see get_layout_info's "
                                + "tables[]), or 'afterControl' on an inline control, instead.");
                        }

                        var newTableGrid = DiffTableGridViolations(baselineTableGrid, TableGridConsistencyGuard.Find(doc));
                        if (newTableGrid.Count > 0)
                        {
                            return ToolResponse.Failure(
                                ErrorCodes.EditWouldBreakTable,
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
                        // through any mutating tool's public surface can actually trigger it today. It is kept as a
                        // backstop against a future change breaking that invariant, and is deliberately not
                        // force-tested — there is no honest way to reach it from the outside without
                        // fabricating an internal bug to order.
                        if (newErrors.Count > 0)
                        {
                            var preview = string.Join(
                                " | ", newErrors.Take(5).Select(e => $"{e.Path?.XPath}: {e.Description}"));
                            return ToolResponse.Failure(
                                ErrorCodes.EditWouldCorrupt,
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
                    var targetDir = Path.GetDirectoryName(commitPath)
                        ?? throw new InvalidOperationException($"Could not determine the directory of '{layoutPath}'.");
                    var staged = Path.Combine(targetDir, $".bcwl-stage-{Guid.NewGuid():N}.docx");
                    try
                    {
                        RetryTransientFileDenial(() => ScratchFiles.CopyWritable(workingCopy, staged));
                        RetryTransientFileDenial(() => File.Move(staged, commitPath, overwrite: true));
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        // Access denied HERE can only be about layoutPath or its folder: the stage copy reads
                        // this call's own working copy, and the move replaces layoutPath. Word holding the file
                        // open, the read-only attribute, and a folder the account cannot write all fail this
                        // way (not as a sharing-violation IOException), so this is the file_locked case - see
                        // TargetFileNotWritableException. layoutPath is unchanged.
                        ScratchFiles.Delete(staged);
                        throw new TargetFileNotWritableException(Path.GetFullPath(layoutPath), ex);
                    }
                    catch
                    {
                        ScratchFiles.Delete(staged);
                        throw;
                    }

                    return onSuccess(mutateResult, quick);
                }
                finally
                {
                    // Best-effort (never throws): see ScratchFiles.Delete for why a cleanup failure must not
                    // replace this call's real outcome.
                    ScratchFiles.Delete(workingCopy);
                }
            }
        }, layoutPath, tool);
    }

    /// <summary>
    /// Name shapes swept by <see cref="SweepStaleStagingFiles"/>: this class's own <c>.bcwl-stage-*.docx</c>
    /// commit-step staging file, <c>BcWordLayout.Merge.MergeEngine</c>'s <c>.bcwl-merge-stage-*.docx</c>
    /// merge-commit staging file, and <see cref="LayoutBuilder"/>'s <c>.bcwl-build-*.docx</c> build file. The
    /// latter two live in DIFFERENT directories in the common case — the merge one in whatever directory a
    /// <c>preview_layout</c>/merge <c>outputDir</c> points at, the build one next to a
    /// <c>create_layout</c> <c>outputPath</c> — so this sweep, which only ever looks at a layout's own
    /// directory, is a complementary pass for them (each one's primary defense is its owner's own sweep,
    /// scoped to the directory it writes into: <c>MergeEngine.SweepStaleMergeStagingFiles</c> and
    /// <see cref="LayoutBuilder.Create"/>'s build-file sweep). Swept here too because the directories often
    /// coincide: layouts are created next to the layouts they are edited beside, and a caller may point a
    /// merge/preview <c>outputDir</c> at a layout's own directory.
    /// </summary>
    private static readonly string[] StagingFileGlobs =
        { ".bcwl-stage-*.docx", ".bcwl-merge-stage-*.docx", ".bcwl-build-*.docx" };

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
    /// SAFE AGAINST A LIVE CONCURRENT COMMIT: only files older than
    /// <see cref="ScratchFiles.OrphanRetentionWindow"/> are removed — comfortably longer than any plausible
    /// in-flight commit (see that field's own remarks, including why creation time is the age signal) — so a
    /// staging file another call (this process, serialized behind the very
    /// <see cref="EditLockFor"/>/<see cref="CrossProcessLock"/> pair guarding this one, or a DIFFERENT
    /// process mutating a DIFFERENT layout that happens to share this directory) is actively mid-commit on
    /// can never be mistaken for an orphan and swept out from under it.
    /// </para>
    /// <para>
    /// BEST-EFFORT: every failure — a locked file (another process's antivirus scan, or the narrow window
    /// where a genuinely live commit is mid-<see cref="File.Move(string, string, bool)"/>), an inaccessible
    /// directory, a candidate deleted by a concurrently-running sweep — is swallowed by
    /// <see cref="ScratchFiles.SweepStale"/>. A missed sweep simply leaves the file for a later mutating call
    /// to retry; it must never fail the edit this cleanup piggybacks on. A READ-ONLY orphan (one a read-only
    /// source left behind before GitHub issue #94 made every private copy writable) is removed too.
    /// </para>
    /// </remarks>
    private static void SweepStaleStagingFiles(string layoutPath)
    {
        string? targetDir;
        try
        {
            targetDir = Path.GetDirectoryName(Path.GetFullPath(layoutPath));
        }
        catch
        {
            return; // Best-effort: a path-shape problem is reported by the edit itself, never by this sweep.
        }

        if (string.IsNullOrEmpty(targetDir))
        {
            return;
        }

        foreach (var glob in StagingFileGlobs)
        {
            ScratchFiles.SweepStale(targetDir, glob);
        }
    }

    /// <summary>The name prefix of <see cref="GuardMutate{TResult}"/>'s <c>%TEMP%</c> edit working copy.</summary>
    private const string EditWorkingCopyPrefix = "bcwl-edit-";

    /// <summary>
    /// How often, at most, one host process sweeps <c>%TEMP%</c> for orphaned edit working copies (see
    /// <see cref="SweepStaleEditWorkingCopiesIfDue"/>). The layout-directory sweep runs on every call because
    /// that directory is small; <c>%TEMP%</c> is shared by every program on the machine and can hold tens of
    /// thousands of entries, and an orphan only needs removing eventually, not on the next call.
    /// </summary>
    private static readonly TimeSpan EditWorkingCopySweepInterval = TimeSpan.FromHours(1);

    /// <summary>UTC ticks before which <see cref="SweepStaleEditWorkingCopiesIfDue"/> skips its sweep.</summary>
    private static long _nextEditWorkingCopySweepTicks;

    /// <summary>
    /// Runs <see cref="SweepStaleEditWorkingCopies"/> against <see cref="Path.GetTempPath"/> when
    /// <see cref="EditWorkingCopySweepInterval"/> has elapsed since this process last did, so the first
    /// mutating call of a fresh process always sweeps. Concurrent callers race on one
    /// <see cref="Interlocked.CompareExchange(ref long, long, long)"/>, so exactly one of them sweeps per interval.
    /// </summary>
    private static void SweepStaleEditWorkingCopiesIfDue()
    {
        var nowTicks = DateTime.UtcNow.Ticks;
        var dueTicks = Interlocked.Read(ref _nextEditWorkingCopySweepTicks);
        if (nowTicks < dueTicks
            || Interlocked.CompareExchange(
                ref _nextEditWorkingCopySweepTicks, nowTicks + EditWorkingCopySweepInterval.Ticks, dueTicks) != dueTicks)
        {
            return;
        }

        SweepStaleEditWorkingCopies(Path.GetTempPath());
    }

    /// <summary>
    /// Best-effort, age-based cleanup of orphaned <c>bcwl-edit-*.docx</c> working copies in
    /// <paramref name="tempDir"/>. <see cref="GuardMutate{TResult}"/> creates one per call in
    /// <see cref="Path.GetTempPath"/> and deletes it in its own <c>finally</c>; a hard-killed process skips
    /// that, and before GitHub issue #94 a read-only source also left a read-only copy that the delete could
    /// not remove. Nothing else ever revisited the family. Same age rule and best-effort contract as
    /// <see cref="SweepStaleStagingFiles"/> (see <see cref="ScratchFiles.SweepStale"/>).
    /// </summary>
    /// <remarks>
    /// <c>%TEMP%</c> is NOT tool-owned, unlike a layout's directory, where the dot-prefixed staging names are
    /// distinctive. So a candidate must also match the exact shape this class generates
    /// (<see cref="IsEditWorkingCopyName"/>: the prefix, a 32-character lowercase hex GUID, <c>.docx</c>), the
    /// same restriction <c>LifecycleTools.SweepStalePreviewDirs</c> applies to its own shared root. Internal
    /// rather than private, and taking the directory as a parameter, so a test can drive it against a private
    /// folder instead of the machine's real <c>%TEMP%</c>.
    /// </remarks>
    internal static void SweepStaleEditWorkingCopies(string tempDir) =>
        ScratchFiles.SweepStale(tempDir, EditWorkingCopyPrefix + "*.docx", IsEditWorkingCopyName);

    /// <summary>True for exactly <c>bcwl-edit-&lt;32 lowercase hex&gt;.docx</c>, the name this class generates.</summary>
    private static bool IsEditWorkingCopyName(string fileName)
    {
        const int GuidLength = 32;
        const string Extension = ".docx";
        if (fileName.Length != EditWorkingCopyPrefix.Length + GuidLength + Extension.Length
            || !fileName.StartsWith(EditWorkingCopyPrefix, StringComparison.Ordinal)
            || !fileName.EndsWith(Extension, StringComparison.Ordinal))
        {
            return false;
        }

        for (var i = EditWorkingCopyPrefix.Length; i < EditWorkingCopyPrefix.Length + GuidLength; i++)
        {
            if (fileName[i] is not ((>= '0' and <= '9') or (>= 'a' and <= 'f')))
            {
                return false;
            }
        }

        return true;
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

    /// <summary>Thin <see cref="GuardMutate{TResult}"/> instantiation building the content-control edit tools' <see cref="EditResultDto"/> — used by <see cref="EditTools"/>'s <c>insert_field</c>/<c>insert_label</c>/<c>insert_text</c>/<c>insert_page_number</c>/<c>insert_picture</c>/<c>remove_control</c> and <see cref="TableTools"/>'s <c>insert_repeater_table</c>/<c>insert_repeater_row</c>/<c>insert_subtotal_row</c>.</summary>
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
    /// tools (<c>insert_table</c>, <c>insert_table_row</c>, <c>set_column_widths</c>, <c>insert_column</c>,
    /// <c>remove_column</c>, <c>merge_cells</c>, <c>split_cells</c>, <c>set_cell_borders</c>), building
    /// <see cref="TableEditResultDto"/>. Same save-or-reject safety and
    /// post-edit QuickValidation guarantee as <see cref="GuardEdit"/>/<see cref="GuardCellEdit"/>, plus the
    /// <see cref="TableGridConsistencyGuard"/> backstop <see cref="GuardMutate{TResult}"/> now runs for every
    /// mutating tool — so a structural table edit that would desync a row's cells from the grid is rejected
    /// with the file untouched. <paramref name="tool"/> is passed by a tool whose parameter names collide with
    /// another tool's, so its <c>invalid_argument</c> hints key on both (see <see cref="ToolSpecificInvalidArgumentHint"/>).
    /// </summary>
    internal static ToolResponse GuardTableEdit(
        string layoutPath, Func<WordprocessingDocument, TableEditResult> mutate, string? tool = null) =>
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
        }, tool);

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
                result.RepointedForeignBindings
                    .Select(r => new RepointedBindingDto(r.Alias, r.XPath, r.Part, r.SdtId, r.PreviousNamespace))
                    .ToList(),
                result.NewUnboundFields,
                new ValidationSummaryDto(
                    result.QuickValidation.Level,
                    result.QuickValidation.Passed,
                    result.QuickValidation.ErrorCount,
                    result.QuickValidation.WarningCount),
                result.RepairedUnreadablePart);

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
            ErrorCodes.FileLocked,
            $"Timed out after {timeout.TotalSeconds:0.#}s waiting for the edit lock on "
            + $"'{Path.GetFullPath(layoutPath)}'.",
            "Another process (e.g. another IDE window's MCP host) is currently editing or previewing this "
            + "layout - wait for it to finish and retry. If nothing else is actually using this file, a prior "
            + "host process may have exited while holding the lock; it is released automatically, so a retry "
            + "after a short pause should succeed.");

    // ---- location building (shared by every tool with a locationType: EditTools' inserts, TableTools' insert_repeater_table/insert_table) ----

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
    /// <para>
    /// Both string parameters are nullable here even though the tools declare them non-null: the SDK binds an
    /// explicit JSON <c>null</c> straight through. A null <paramref name="layoutPart"/> means its documented
    /// default (<c>body</c>), like omitting it; a null <paramref name="locationType"/> (required, no default)
    /// is refused as <c>invalid_argument</c> rather than crashing as a <see cref="NullReferenceException"/>.
    /// </para>
    /// </summary>
    internal static Location BuildLocation(
        string? locationType, int? controlId, int? tableIndex, int? row, int? col, string? searchText,
        string? layoutPart = "body", string? partName = null)
    {
        if (!TryParseLocationKind(locationType, out var kind))
        {
            throw new ArgumentException(
                (locationType is null ? "locationType is required (got null)." : $"Unknown locationType '{locationType}'.")
                + " Use 'documentEnd', 'afterControl', 'tableCell', or 'atText'.",
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

    private static bool TryParseLocationKind(string? locationType, out LocationKind kind)
    {
        switch (locationType?.Trim().ToLowerInvariant())
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
    /// <see cref="LayoutPart"/> rather than building a full <see cref="Location"/>). A null
    /// <paramref name="layoutPart"/> — an explicit JSON <c>null</c>, which the SDK binds straight through — is
    /// the documented default, <c>body</c>, exactly as if the argument had been omitted.
    /// </summary>
    internal static bool TryParseLayoutPart(string? layoutPart, out LayoutPart part)
    {
        switch ((layoutPart ?? "body").Trim().ToLowerInvariant())
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
                    cell.InnerControlIds,
                    cell.GridColumn,
                    cell.GridSpan)).ToList(),
                r.GridBefore,
                r.GridAfter,
                r.RepeaterControlId)).ToList());

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
    /// A caller file that is not the kind of file its parameter takes - not a <c>.docx</c> package at all, or
    /// not well-formed XML (GitHub issue #95) - reaches <c>invalid_layout</c> through
    /// <see cref="InvalidInputFileException"/>, which the throw site that knows the file wraps around the BCL's
    /// <see cref="FileFormatException"/>/<see cref="XmlException"/>, so the hint can name that file and the
    /// line/position. A bare <see cref="XmlException"/> is caught too, as the fallback for a damaged part the
    /// SDK parses lazily (see that branch). <see cref="FileFormatException"/> itself is deliberately NOT
    /// mapped: <see cref="LayoutBuilder"/>'s internal-bug gate has its own type now
    /// (<see cref="BuildSelfCheckFailedException"/>), but a stray one from anywhere else is still not
    /// evidence the caller's input was wrong. A <c>dataOverridesPath</c> file that cannot be used is an
    /// <see cref="ArgumentException"/> for that parameter instead, so it reports <c>invalid_argument</c>.
    /// </para>
    /// <para>
    /// <c>file_locked</c> is produced three ways: (1) below, by catching
    /// <see cref="IOException"/> and checking <see cref="Exception.HResult"/> for a Windows sharing/lock
    /// violation (<c>ERROR_SHARING_VIOLATION</c> 0x80070020 / <c>ERROR_LOCK_VIOLATION</c> 0x80070021) - the
    /// "this layout is open in Word (or another program) right now" case, where opening the file itself
    /// fails; (2) below, by catching <see cref="TargetFileNotWritableException"/> - the commit that REPLACES
    /// the caller's target was refused with access denied, which is how Word holding the file open, the
    /// read-only attribute and a missing write permission all surface at a rename (GitHub issue #94; reads
    /// and copies of a file Word holds still succeed, so the denial only shows at the commit); and (3) not
    /// through this method at all, but as a plain <see cref="ToolResponse.Failure"/>
    /// returned directly by <see cref="CrossProcessLockTimeoutFailure"/> when the cross-process/in-process
    /// edit-lock WAIT itself times out (see <see cref="GuardMutate{TResult}"/>, <see cref="GuardRead"/>, and
    /// <c>LifecycleTools.PreviewLayout</c>/<c>CreateLayout</c>) - the "another process/thread is currently
    /// editing this same layout" case, which is not an exception at all and so never reaches this catch chain.
    /// All land on the identical <c>file_locked</c> code so a caller need only branch on one value regardless
    /// of which actually happened. An <see cref="IOException"/> whose <see cref="Exception.HResult"/>
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
    /// <para>
    /// A <see cref="LayoutDefectException"/> - the package carries a defect no argument can get past, e.g. a
    /// BC dataset part that is there but does not parse - is <c>invalid_layout</c> too, with a hint chosen by
    /// its typed <see cref="LayoutDefectException.Defect"/> (<see cref="LayoutDefectHintFor"/>) that names the
    /// repair, rather than the generic "confirm the file contains a dataset part" one.
    /// </para>
    /// </remarks>
    internal static ToolResponse Guard(Func<ToolResponse> body, string path, string? tool = null)
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
                ErrorCodes.FileNotFound,
                $"{ex.Message} Path checked: '{missingPath}'.",
                "Pass an absolute path (not relative) to a file that already exists on disk - the parameter "
                + "named above must point at it; check for typos, a missing drive letter, or a relative path "
                + "that does not resolve the way you expect.");
        }
        catch (IOException ex) when (IsSharingOrLockViolation(ex))
        {
            // "This layout is open in Word (or another program) right now" - the
            // sharing-violation part of the file_locked story (see this method's own remarks above for the
            // other two - a refused write of the target, and a timed-out edit-lock WAIT).
            return ToolResponse.Failure(
                ErrorCodes.FileLocked,
                ex.Message,
                "This file is open in another program (typically Word) or is being read/written by another "
                + "process right now - close it there (or wait for that other operation to finish) and retry; "
                + "the layout on disk is unaffected by this failure.");
        }
        catch (TargetFileNotWritableException ex)
        {
            // The commit that replaces the caller's target file was refused with access denied (GitHub issue
            // #94) - see the exception type's own remarks for why each of the three causes below fails this
            // way rather than as a sharing violation. Keyed on the TYPE, which only a throw site writing the
            // target constructs; a bare UnauthorizedAccessException (e.g. a source file that cannot be read)
            // still reports internal_error.
            return ToolResponse.Failure(
                ErrorCodes.FileLocked,
                ex.Message,
                $"'{ex.FilePath}' could not be replaced. The likely causes, most common first: (1) it is open "
                + "in Word or another program - close it there and retry; (2) it is read-only, e.g. a "
                + "TFVC/Perforce workspace file that is not checked out - check it out or clear its read-only "
                + "attribute, then retry; (3) your account has no write permission on the file or its folder - "
                + "fix the permission, or write to a folder you own. The file itself is unchanged.");
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
                ErrorCodes.InvalidLayout,
                ex.Message,
                "This file exceeds one of this tool's supported size/nesting limits (a package part or "
                + "schema file too large, too many parts, or schema/document nesting too deep) - "
                + "it is likely malformed or a maliciously crafted file rather than a real Business Central "
                + "layout/schema. This is not something to retry with the same file.");
        }
        catch (InvalidInputFileException ex)
        {
            // A caller's layout/template/schema file that is not a .docx package or not well-formed XML
            // (GitHub issue #95) - wrapped at the throw site that knows WHICH file it was reading, so the hint
            // can name it (see the exception type's own remarks for why the BCL types are not mapped wholesale).
            return ToolResponse.Failure(
                ErrorCodes.InvalidLayout,
                ex.Message,
                InvalidInputFileHint(ex.Problem, ex.FilePath, ex.LineNumber, ex.LinePosition));
        }
        catch (XmlException ex)
        {
            // Fallback for the one XmlException source no throw site can wrap: the OpenXML SDK parses a
            // package's parts LAZILY, at first access anywhere in a tool body, so a damaged part inside an
            // otherwise valid .docx surfaces here. Every XML this codebase reads comes from a caller's file
            // (dataset paths are evaluated as XPath, which throws XPathException, not this type), and every
            // tool's lazily-parsed package is the file `path` names - an edit's or preview's private copy is
            // a copy OF it, and create_layout's template is wrapped at its own source.
            return ToolResponse.Failure(
                ErrorCodes.InvalidLayout,
                $"'{path}' contains an XML part that is not well-formed: {ex.Message}",
                InvalidInputFileHint(
                    InputFileProblem.MalformedPackagePart,
                    path,
                    ex.LineNumber > 0 ? ex.LineNumber : null,
                    ex.LineNumber > 0 ? ex.LinePosition : null));
        }
        catch (InvalidDataException ex)
        {
            return ToolResponse.Failure(
                ErrorCodes.InvalidLayout,
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
                ErrorCodes.TemplateNotUnbound,
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
        catch (LayoutDefectException ex)
        {
            // The package itself carries a defect this call cannot work around (see LayoutDefectException's
            // own remarks for why that is not just InvalidDataException's generic invalid_layout hint):
            // keyed off the typed Defect, never ex.Message, exactly like NotFoundException's TargetKind.
            return ToolResponse.Failure(ErrorCodes.InvalidLayout, ex.Message, LayoutDefectHintFor(ex.Defect));
        }
        catch (ArgumentException ex)
        {
            return ToolResponse.Failure(
                ErrorCodes.InvalidArgument,
                ex.Message,
                ToolSpecificInvalidArgumentHint(tool, ex.ParamName) ?? InvalidArgumentHint(ex.ParamName));
        }
        catch (NotFoundException ex)
        {
            return ToolResponse.Failure(ErrorCodes.NotFound, ex.Message, NotFoundHintFor(ex.TargetKind));
        }
        catch (Exception ex)
        {
            return ToolResponse.Failure(
                ErrorCodes.InternalError,
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
    /// Builds the <c>invalid_layout</c> hint for a caller file that could not be read as what its parameter
    /// takes (see <see cref="InputFileProblem"/>), naming <paramref name="filePath"/> and, for XML, the line
    /// and position the reader stopped at. Keyed off the typed problem, never the exception's message text.
    /// </summary>
    /// <remarks>
    /// The <see cref="InputFileProblem.NotAWordPackage"/> hint explains the <c>.docx</c>-versus-schema-<c>.xml</c>
    /// split because that is the usual mistake behind it: every tool that takes either kind of file decides by
    /// EXTENSION (a path ending in <c>.docx</c> is opened as a package, anything else is read as a schema
    /// <c>.xml</c>), so a schema saved under a <c>.docx</c> name, or a layout under any other name, lands on
    /// the wrong reader. The <see cref="InputFileProblem.MalformedXmlFile"/> hint states the reverse case.
    /// </remarks>
    private static string InvalidInputFileHint(InputFileProblem problem, string filePath, int? line, int? position)
    {
        var where = line is > 0 ? $" (line {line}, position {position})" : "";
        return problem switch
        {
            InputFileProblem.NotAWordPackage =>
                $"'{filePath}' is not a Word .docx package, or is too damaged to open. A .docx (or .dotx) is a "
                + "zip container, so a schema .xml or a text file renamed to .docx, an empty (0-byte) file, or a "
                + "truncated copy or download cannot be opened as one. Pass the real BC Word layout (saved by Word "
                + "or exported from Business Central); if Word offers to repair this file, save the repaired copy "
                + "and pass that. A standalone exported schema .xml is accepted only as create_layout's "
                + "schemaSource, refresh_xml_part's newSchemaSource or list_dataset_fields' source, and only "
                + "under its .xml extension: a path ending in .docx is always opened as a layout package.",

            InputFileProblem.MalformedXmlFile =>
                $"'{filePath}' is not well-formed XML{where}, so it cannot be read as a report dataset schema. "
                + "Typical causes: a truncated copy, an encoding declaration that does not match the bytes (e.g. "
                + "encoding=\"utf-16\" on a file saved as UTF-8), or a DOCTYPE/DTD, which is refused for "
                + "security. Re-export the schema .xml from Business Central rather than editing it by hand. A "
                + "path that does not end in .docx is always read as a schema .xml, so if this file is really a "
                + ".docx layout, give it its .docx extension.",

            _ =>
                $"'{filePath}' opened as a Word package, but an XML part inside it is not well-formed{where}: the "
                + "package is damaged (e.g. a part was hand-edited or truncated). Open it in Word - if Word "
                + "repairs it, save the repaired copy and retry with that; otherwise restore the file from "
                + "source control.",
        };
    }

    /// <summary>
    /// The <c>invalid_argument</c> hint for a parameter whose NAME means different things on different tools,
    /// keyed by tool + parameter; null when the pair has no tool-specific hint, so <see cref="InvalidArgumentHint"/>'s
    /// name-only map applies. It exists because a name-only key cannot tell <c>insert_table</c>'s integer
    /// <c>columns</c> (a column COUNT) from <c>insert_repeater_table</c>'s comma-separated list of column NAMES,
    /// or <c>insert_table</c>'s <c>rows</c> from <c>preview_layout</c>'s sample-row count — the hint used to
    /// describe the wrong tool's parameter (GitHub issue #98). The tool name arrives through
    /// <see cref="Guard"/>'s <c>tool</c> argument, which only the tools that own one of these colliding names
    /// pass; every other call keeps the name-only hint. Each pair is pinned through the real server pipeline by
    /// <c>McpHostPipelineTests</c>.
    /// </summary>
    private static string? ToolSpecificInvalidArgumentHint(string? tool, string? paramName) =>
        (tool, paramName?.ToLowerInvariant()) switch
        {
            ("insert_table", "rows") =>
                "insert_table's rows is how many rows the new table has: an integer from 1 to 100.",

            ("insert_table", "columns") =>
                "insert_table's columns is how many columns the new table has: an integer from 1 to 30 (not a "
                + "list of column names - that is insert_repeater_table's columns). Size them with columnWidths, "
                + "one width in twips per column.",

            ("preview_layout", "rows") =>
                "preview_layout's rows is how many sample rows to generate per repeating section: 1 or more "
                + "(default 3). Values above the internal cap of 100 are capped with a row-cap warning. To preview "
                + "real records instead of sample rows, pass an exported dataset via dataOverridesPath.",

            _ => null,
        };

    /// <summary>
    /// Builds an <c>invalid_argument</c> hint tailored to the specific argument that failed, keyed off
    /// <see cref="ArgumentException.ParamName"/> (populated via <c>nameof(...)</c> at every throw site across
    /// this tool surface - <c>ToolGuards</c>/the tool-family classes themselves, <c>SdtFactory</c>,
    /// <c>LayoutEditor</c>, <c>Location.Validate</c>, <c>LocationResolver</c>, <c>TableStructureEditor</c>,
    /// <c>CellTextEditor</c>, <c>LayoutBuilder</c>, <c>LayoutRefresher</c>, and <c>SampleDataGenerator</c>): names that argument's valid
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
            + "default) is supported - a repeater TABLE in a header/footer is not supported, since nothing "
            + "shows that Business Central renders a repeating section there (unlike insert_field/insert_label, "
            + "which do support 'header'/'footer'). For the insert tools: a "
            + "location that resolves INSIDE a data-bound field/label/picture control's own content (afterControl "
            + "on a cell-level control, tableCell on a cell that IS such a control, atText matching its text) is "
            + "refused, because Word discards or rejects anything placed there; use tableCell on a plain cell "
            + "(get_layout_info's tables[] shows which cells are controls) or afterControl on an inline control "
            + "instead.",

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
            + "insert_repeater_table's and wrap_in_repeater's dataItem must be a repeating, non-system DATA ITEM "
            + "(not a leaf column); wrap_in_repeater also needs at least one control in its range bound under "
            + "that item, and no repeater in the range over data outside it.",

        "columns" or "column" =>
            "insert_repeater_table's columns must be a non-empty, comma-separated list of LEAF column names "
            + "belonging to the given dataItem (not full paths, not nested data items) - see "
            + "list_dataset_fields for the exact column names available under that data item.",

        "columnwidths" or "options" =>
            "columnWidths, if supplied, must be a comma-separated list of integers with exactly one width "
            + $"(in twips, each 1 to {TableTools.MaxColumnWidthTwips}) per column, matching the column count "
            + "exactly - omit it entirely for an even default width per column instead.",

        "widths" =>
            "set_column_widths' widths must be a non-empty comma-separated list of integers (twips, each 1 to "
            + $"{TableTools.MaxColumnWidthTwips}), exactly one per GRID column - the count must equal the table's "
            + "columnCount from get_layout_info (a cell's own columns are positions "
            + "gridColumn..gridColumn+gridSpan-1 in that list).",

        "cells" =>
            "cells is a comma-separated list of cell specs, one per cell of the new row: '-' is an empty spacer, "
            + "a name is a bound cell (insert_repeater_row: a leaf column of the child data item; "
            + "insert_table_row/insert_subtotal_row: a FULL dataset path), 'a+b' chains several in one cell, and "
            + $"an optional 'N:' prefix spans N grid columns (1 to {TableTools.MaxCellSpan}). The spans must sum to "
            + "the table's columnCount (get_layout_info's tables[]), e.g. '4:-,3:/Header/Totals/TotalText,"
            + "/Header/Totals/TotalAmount' on an 8-column table.",

        "atcolumn" =>
            "insert_column's atColumn is the 0-based GRID position for the new column (0..columnCount from "
            + "get_layout_info; columnCount, or omitting it, appends at the far-right edge): a cell's "
            + "gridColumn in get_layout_info's tables[] inserts before that cell, gridColumn+gridSpan after "
            + "it - not its colIndex, the physical position in the row. An INTERIOR "
            + "position must land on a cell boundary in every row that carries the new column's content (a "
            + "header row, and the bound data row) - if it falls inside a spanned cell there, split that "
            + "cell first with split_cells or pick a different position.",

        "width" =>
            "insert_column's width is the NEW column's width in twips: a positive number smaller than the "
            + "table's current total width (the sum of get_layout_info's gridColumnWidths), because the table "
            + "keeps that total and the new column's twips are taken proportionally from the existing "
            + "columns. Omit it for an equal share, and use set_column_widths afterwards for any other split.",

        "gridcolumn" =>
            "remove_column's column is a 0-based GRID column index (0..columnCount-1 from get_layout_info) - "
            + "the gridColumn get_layout_info's tables[] report for the column's cells, not their colIndex; "
            + "the last remaining column cannot be removed, nor a column that is some row's ONLY cell (a row "
            + "otherwise made of w:gridBefore/w:gridAfter skipped columns - Word cannot open a row with no "
            + "cells), nor any column of a table none of whose rows matches its w:tblGrid (repair the grid "
            + "first). The message above names the row or the mismatch.",

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

        "format" =>
            "format is one object of optional members: bold, italic and underline (true or false), "
            + "fontSizePoints (4 to 96, halves allowed), fontName (a font family name such as 'Segoe UI'), color "
            + "(six hex digits RRGGBB such as '0070C0', or 'auto') and characterStyle (the id of a CHARACTER style "
            + "the layout's styles part defines - the message above lists them). Omit a member to leave that "
            + "property as it is.",

        "alignment" =>
            "alignment must be 'left', 'center' or 'right' (case-insensitive); omit it to keep the paragraph's "
            + "existing alignment.",

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
            + "reverse. In a line-items table the field is bound only into the rows of the repeater whose data "
            + "item contains dataPath (get_layout_info's controls[] show each repeater's xpath), so it must be "
            + "a column of one of the table's repeated data items - or of a parent item when the table has a "
            + "single repeater.",

        "headertext" or "headerlabelpath" =>
            "insert_column's headerText/headerLabelPath fill the new column's cell in the table's HEADER "
            + "row(s): rows marked w:tblHeader, else the first (non-repeater) row of a line-items table, else - "
            + "in a plain table - its caption rows (rows whose bound controls are all labels). A table with "
            + "none has nowhere to put them: omit them and add the caption afterwards with set_cell_text or "
            + "insert_label at the cell you mean. headerLabelPath must also be a label-shaped column exactly as "
            + "list_dataset_fields reports it (isLabel=true).",

        "keeptext" =>
            "remove_control's keepText=true cannot target a ROW-level repeating-section control (it would orphan "
            + "its row template from any enclosing repeater), nor the w:group control between a block-level "
            + "top-level repeater and its item (BC would print an empty document); to remove the repeater "
            + "with its content, call remove_control with keepText=false (the default) on the REPEATER's id, "
            + "otherwise leave the control in place. (keepText=true on a BLOCK-level repeater's own id does "
            + "work: it unwraps the whole repeater and keeps its content.)",

        "firstcontrolid" or "lastcontrolid" or "wholebody" =>
            "wrap_in_repeater repeats the body blocks holding firstControlId through lastControlId: control ids "
            + "from get_layout_info's controls[] (sdtId), in the DOCUMENT BODY, the first one earlier in the "
            + "document than the last (lastControlId defaults to firstControlId), with no section break "
            + "between them. Or pass wholeBody=true and no ids to repeat the whole body.",

        "schema" =>
            "A bound control must name the storeItemID of the layout's dataset part, and every insert tool "
            + "resolves one from the layout itself before building the control (from the part's itemProps, or "
            + "from the storeItemID its existing bindings agree on - refusing with invalid_layout when there is "
            + "neither), so reaching this message indicates a tool bug rather than a mistake in your call; "
            + "please report it with the message above.",

        "schemasource" =>
            "create_layout's schemaSource must be a non-empty absolute path to an existing .docx layout "
            + "(its BC dataset part is reused) or a standalone exported schema .xml; see the message above "
            + "for exactly what was rejected about it.",

        "templatepath" =>
            "create_layout's templatePath must be an unbound branded/styled shell saved as a Word Document "
            + "(.docx) or a Word Template (.dotx - converted to a document automatically). A macro-enabled "
            + "file (.dotm/.docm) is refused rather than stripped of its macros: open it in Word, save it as a "
            + "Word Document or Word Template, and pass that file instead (or omit templatePath for a blank "
            + "layout).",

        "newschemasource" =>
            "refresh_xml_part's newSchemaSource must be a non-empty absolute path to an existing .docx "
            + "layout (its BC dataset part is reused) or a standalone exported schema .xml - the same shape "
            + "create_layout's schemaSource accepts; see the message above for exactly what was rejected.",

        "outputpath" =>
            "The output path must be a non-empty absolute path (create_layout's outputPath is where the new "
            + "layout .docx is written); see the message above for exactly what was rejected about it. It "
            + "must not be schemaSource itself while that file is an authored layout with bound controls: "
            + "create_layout keeps only the dataset part. To bind an existing layout to a new or changed "
            + "schema in place, call refresh_xml_part on it instead (layoutPath = that layout, newSchemaSource "
            + "= the new schema); to start a new layout from its schema, pass a different outputPath.",

        "dataoverridespath" =>
            "preview_layout's dataOverridesPath must be a well-formed XML file in one of the two shapes Business "
            + "Central produces for THIS layout's report: the layout's own data-store part (root "
            + "'NavWordReportXmlPart' in the 'urn:microsoft-dynamics-nav/reports/<ReportName>/<id>/' namespace "
            + "get_layout_info reports), or the report's Send to > XML export (root 'ReportDataSet', whose id "
            + "must be this layout's report id). See the message above for the file and what was rejected; "
            + "omit dataOverridesPath to preview with generated sample data instead.",

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

    /// <summary>
    /// Builds the <c>invalid_layout</c> hint for a <see cref="LayoutDefectException"/>, keyed off its typed
    /// <see cref="LayoutDefectException.Defect"/> (the <see cref="NotFoundHintFor"/> pattern): each names the
    /// repair the PACKAGE needs, since no change to the call's own arguments can get past these.
    /// </summary>
    private static string LayoutDefectHintFor(LayoutDefect defect) => defect switch
    {
        LayoutDefect.UnreadableDatasetPart =>
            "The layout's BC dataset part is present but is not well-formed XML (the message names the part and "
            + "the parser's error) - typically an encoding declaration that no longer matches its bytes after a "
            + "re-save in a text editor, an unescaped '&', or a truncated file. refresh_xml_part repairs it in "
            + "place: pass this layout as layoutPath and an intact copy of the same report's layout (or its "
            + "exported schema .xml) as newSchemaSource; the part keeps its storeItemID, so every binding stays "
            + "linked. If the broken file was passed as a schema SOURCE (schemaSource/newSchemaSource), pass an "
            + "intact copy instead.",

        LayoutDefect.SeveralDatasetParts =>
            "A layout carries exactly one BC dataset part; this one carries several (the message lists them and "
            + "names the one its bindings use - normally the one to keep). Remove the extra part from the .docx "
            + "package: delete its customXml/itemN.xml, the matching customXml/itemPropsN.xml and "
            + "customXml/_rels/itemN.xml.rels, its relationship in word/_rels/document.xml.rels and its overrides "
            + "in [Content_Types].xml, then retry. get_layout_info and validate_layout still read the file (its "
            + "single-bc-part finding lists the parts too). If the file was passed as a schema SOURCE "
            + "(schemaSource/newSchemaSource), pass a layout with a single dataset part instead.",

        LayoutDefect.DuplicateControlId =>
            "Content-control ids must be unique for id addressing to mean anything (validate_layout's "
            + "duplicate-control-id warning lists every shared id). No tool here renumbers a control: for an "
            + "insert, address the spot another way (tableCell or atText instead of afterControl); to fix the "
            + "package, give all but one of the controls the message names a new w:sdtPr/w:id value (any integer "
            + "no other control uses), or open the layout in Word and save it - the spec (ECMA-376 Part 1, "
            + "17.5.2.18) has a consumer keep the first and assign fresh ids to later duplicates when it loads the "
            + "document. Then call get_layout_info for the ids as they now are.",

        LayoutDefect.NoStoreItemId =>
            "A new bound control must name the storeItemID of the layout's dataset part. When that part carries "
            + "no itemProps (stock layouts such as StandardStatement ship that way) the tools reuse the "
            + "storeItemID every existing dataset binding names - and this layout offers neither. Business "
            + "Central re-attaches the store item when a layout is uploaded (see validate_layout's single-bc-part "
            + "warning), so one way out is to upload the layout to Business Central and export it back, then "
            + "edit the exported copy, whose dataset part should carry one; another is to work on a copy of the "
            + "layout whose existing bindings all name the same storeItemID. A new itemProps is deliberately not "
            + "invented here: every existing binding would then name a store item the package does not have.",

        _ =>
            "The layout package carries a defect this call cannot work around (see the message above); run "
            + "validate_layout on it for the full list of findings and their repairs.",
    };
}
