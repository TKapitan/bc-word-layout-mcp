namespace BcWordLayout.Domain;

/// <summary>
/// The file-system hygiene shared by every PRIVATE copy this codebase makes of a caller's file: the
/// <c>%TEMP%\bcwl-edit-*</c> edit working copy and the <c>.bcwl-stage-*</c> commit stage
/// (<c>BcWordLayout.McpHost.Tools.ToolGuards</c>), the <c>.bcwl-merge-stage-*</c> merge stage
/// (<c>BcWordLayout.Merge.MergeEngine</c>), and the <c>.bcwl-build-*</c> build file (<see cref="LayoutBuilder"/>).
/// It makes a copy writable, deletes one without letting the cleanup replace the call's real outcome, and
/// sweeps the orphans a killed process leaves behind. Each call site still owns WHICH directory it sweeps
/// and why; only the mechanics live here, so every family follows the same age and read-only rules.
/// </summary>
/// <remarks>
/// WHY THE READ-ONLY ATTRIBUTE IS CLEARED. <see cref="File.Copy(string, string, bool)"/> carries the
/// source's attributes over, <see cref="FileAttributes.ReadOnly"/> included, and a read-only source is the
/// norm in a TFVC/Perforce workspace until the file is checked out. A read-only private copy broke every
/// step after the copy: opening it for write (the edit working copy, the merge stage, a template's build
/// file) failed with an access-denied <c>internal_error</c> naming a temp file the caller never passed, and
/// <see cref="File.Delete"/> refuses a read-only file, so the copy also survived its own cleanup and every
/// sweep (GitHub issue #94). Only the private copy loses the attribute; the caller's own file keeps it. A
/// caller file that is itself read-only is still refused where it would be replaced (see
/// <see cref="TargetFileNotWritableException"/>).
/// </remarks>
public static class ScratchFiles
{
    /// <summary>
    /// How long a private copy must have existed before <see cref="SweepStale"/> treats it as orphaned. A
    /// copy is normally consumed or deleted within the call that made it, so a genuine orphan is either brand
    /// new (another call's live copy, in this process or another one) or untouched for a long time (a
    /// hard-killed process never reached its cleanup). One day comfortably clears any plausible in-flight
    /// window (network-share latency, an antivirus scan holding the handle, a slow disk) while still being
    /// "soon" on human timescales, so an orphan does not linger next to a source-controlled layout.
    /// </summary>
    /// <remarks>
    /// AGE SIGNAL: <see cref="File.GetCreationTimeUtc(string)"/>, deliberately NOT
    /// <see cref="File.GetLastWriteTimeUtc(string)"/>. <see cref="File.Copy(string, string, bool)"/> copies the
    /// SOURCE's last-write time onto the new file but resets its creation time to "now" (confirmed on
    /// Windows/NTFS). Every family here is produced by copying a caller's file, which is often an old,
    /// checked-in layout, so a copy created this instant can already look a day stale by last-write time.
    /// That exact failure once let a concurrent sweep delete a live, mid-commit staging file. Creation time
    /// reflects how long THIS copy has existed, whatever its content's history.
    /// </remarks>
    public static readonly TimeSpan OrphanRetentionWindow = TimeSpan.FromDays(1);

    /// <summary>
    /// Copies <paramref name="sourcePath"/> to <paramref name="destinationPath"/> (overwriting) and clears
    /// <see cref="FileAttributes.ReadOnly"/> on the copy, so a read-only source still yields a private copy
    /// that can be opened for write and deleted. Callers that need the transient-denial retry wrap this call
    /// in <see cref="TransientFileRetry.Run"/>, exactly as they wrapped the bare copy.
    /// </summary>
    public static void CopyWritable(string sourcePath, string destinationPath)
    {
        File.Copy(sourcePath, destinationPath, overwrite: true);
        ClearReadOnly(destinationPath);
    }

    /// <summary>Clears <see cref="FileAttributes.ReadOnly"/> on <paramref name="path"/> when it is set.</summary>
    public static void ClearReadOnly(string path)
    {
        var attributes = File.GetAttributes(path);
        if ((attributes & FileAttributes.ReadOnly) != 0)
        {
            File.SetAttributes(path, attributes & ~FileAttributes.ReadOnly);
        }
    }

    /// <summary>
    /// Deletes the private copy at <paramref name="path"/> if it exists: clears its read-only attribute
    /// first, and retries a transient denial (an antivirus scan of the freshly written file) the way
    /// <see cref="TransientFileRetry"/> does for the copy itself.
    /// </summary>
    /// <remarks>
    /// BEST-EFFORT, NEVER THROWS an I/O failure. Every caller runs this from a <c>finally</c> block or a
    /// <c>catch</c> that is about to rethrow, so a throwing cleanup would REPLACE the call's real outcome:
    /// an edit that already committed would report failure, and a failed open would be reported as a failed
    /// delete of a temp file the caller never heard of. A copy that survives is an orphan that the
    /// age-based <see cref="SweepStale"/> pass for its family removes on a later call.
    /// </remarks>
    public static void Delete(string path)
    {
        try
        {
            if (!File.Exists(path))
            {
                return;
            }

            TransientFileRetry.Run(() =>
            {
                ClearReadOnly(path);
                File.Delete(path);
            });
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // Best-effort: left for a later SweepStale pass (see remarks).
        }
    }

    /// <summary>
    /// Best-effort, age-based removal of orphaned private copies: deletes every file in
    /// <paramref name="directory"/> (not its subdirectories) that matches <paramref name="searchPattern"/>,
    /// passes <paramref name="isOwnedName"/> when one is given, and was created more than
    /// <see cref="OrphanRetentionWindow"/> ago (read-only or not). A fresh match is left alone, since it may be
    /// another call's live copy.
    /// </summary>
    /// <param name="isOwnedName">
    /// Optional exact-shape check on the file NAME, for a directory this tool does not own (the shared
    /// <c>%TEMP%</c>), where a glob alone could match a file another program created.
    /// </param>
    /// <remarks>
    /// BEST-EFFORT: every failure is swallowed, whether a locked candidate, an inaccessible directory, or a
    /// candidate another sweep removed first. A missed sweep leaves the file for a later call; it must never
    /// fail the call this cleanup piggybacks on.
    /// </remarks>
    public static void SweepStale(string directory, string searchPattern, Func<string, bool>? isOwnedName = null)
    {
        try
        {
            if (!Directory.Exists(directory))
            {
                return;
            }

            var cutoffUtc = DateTime.UtcNow - OrphanRetentionWindow;
            foreach (var candidate in Directory.EnumerateFiles(directory, searchPattern))
            {
                try
                {
                    if (isOwnedName is not null && !isOwnedName(Path.GetFileName(candidate)))
                    {
                        continue;
                    }

                    // CreationTimeUtc, not LastWriteTimeUtc: see OrphanRetentionWindow's remarks.
                    if (File.GetCreationTimeUtc(candidate) < cutoffUtc)
                    {
                        ClearReadOnly(candidate);
                        File.Delete(candidate);
                    }
                }
                catch
                {
                    // Best-effort: leave this one candidate for a later sweep.
                }
            }
        }
        catch
        {
            // Best-effort: an inaccessible/unreadable directory must never fail the call.
        }
    }
}
