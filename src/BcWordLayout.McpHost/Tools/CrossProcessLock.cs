using System.Security.Cryptography;
using System.Text;

namespace BcWordLayout.McpHost.Tools;

/// <summary>
/// Cross-process counterpart to <see cref="ToolGuards.EditLockFor"/>'s in-process per-path lock: a
/// named, system-wide <see cref="Mutex"/> keyed by a hash of the layout's normalized,
/// case-folded full path, so TWO SEPARATE PROCESSES (e.g. two IDE windows, each hosting its own MCP server)
/// editing/previewing/reading the SAME layout serialize against each other too — not just two threads inside
/// one process, which is all <see cref="ToolGuards.EditLockFor"/>'s plain <see cref="object"/>/<c>lock</c>
/// can ever see.
/// </summary>
/// <remarks>
/// <para>
/// NAMING/SCOPE: <c>Local\bcwl-edit-&lt;24 lowercase hex chars (96 bits) of a SHA-256 over the case-folded,
/// normalized path&gt;</c>. <c>Local\</c> scopes the name to the current Terminal Services session (the calling user's
/// login session) rather than the whole machine (<c>Global\</c>, which can need elevated rights in some
/// configurations) — exactly the right scope for "the same developer running several IDE windows," and it
/// needs no admin rights. Hashing rather than embedding the path itself sidesteps the ~260-character kernel
/// object name limit and the characters a real path can contain that a kernel object name cannot (backslash
/// is the namespace separator there; the raw path would also make a name trivially longer than the limit for
/// deeply nested repos). The path is upper-invariant-folded before hashing for the same reason
/// <see cref="ToolGuards.EditLockFor"/>'s dictionary is <see cref="StringComparer.OrdinalIgnoreCase"/> and
/// <c>LifecycleTools.PreviewPathHash</c> folds case too: Windows paths are case-insensitive, so two
/// differently-cased spellings of the same file must resolve to the SAME mutex, or the two primitives could
/// disagree about which callers are actually contending for the same file.
/// </para>
/// <para>
/// ACQUIRE ORDER: every caller takes the in-process lock (<see cref="ToolGuards.EditLockFor"/>) FIRST, then
/// this mutex INSIDE it — never the reverse, and never this mutex alone. The in-process lock is cheap and
/// serializes same-process callers (the common case — most calls in a session target paths only this one
/// host process ever touches) without ever leaving managed code; the mutex is the more expensive, kernel-
/// level primitive, needed only for the rarer second-host-process case. Taking the mutex INSIDE the in-process
/// lock means at most ONE thread per process ever waits on a given path's mutex at a time — every caller in
/// every tool family (<see cref="ToolGuards.GuardMutate{TResult}"/>, <c>LifecycleTools.PreviewLayout</c>/
/// <c>CreateLayout</c>, <see cref="ToolGuards.GuardRead"/>) uses this exact same order, so there is no lock-
/// ordering deadlock risk between the two primitives: a deadlock needs two threads each holding one resource
/// while waiting on the other in OPPOSITE orders, and that second order never occurs anywhere in this
/// codebase.
/// </para>
/// <para>
/// ABANDONED MUTEX: <see cref="AbandonedMutexException"/> means the PREVIOUS holder's PROCESS exited (crash,
/// kill, host restart) while still owning the mutex — .NET surfaces this as an exception on the very wait
/// that successfully acquires it next, not as a silent "someone else's ownership continues" state. Treated
/// here as an ordinary successful acquisition (not rethrown, not treated as a fresh timeout): the layout file
/// itself can never be left half-written by that dead process, because every mutating tool's actual commit is
/// the atomic, same-volume <see cref="File.Move(string, string, bool)"/> in
/// <see cref="ToolGuards.GuardMutate{TResult}"/> (mirrored by <c>LayoutBuilder.Create</c>'s own build-then-
/// rename for <c>create_layout</c>) — either that rename already completed (the file is the fully committed
/// NEW version) or it did not (the file is still the fully intact PREVIOUS version); there is no partially-
/// committed state on disk for the abandoned mutex to have been protecting that this caller could now
/// observe. Proceeding as the new owner is therefore safe, not merely convenient.
/// </para>
/// </remarks>
internal static class CrossProcessLock
{
    /// <summary>
    /// Acquire timeout for every MUTATING tool (via <see cref="ToolGuards.GuardMutate{TResult}"/> and
    /// <c>LifecycleTools.CreateLayout</c>) and for <c>preview_layout</c> — all of these already hold the
    /// in-process lock for their whole duration, so this is the bound on how long one process will wait for
    /// ANOTHER process's edit/preview/create of the same path to finish before giving up with
    /// <c>file_locked</c>. Mutable (not <c>const</c>/<c>readonly</c>) so a test can substitute a short value —
    /// any test that does so MUST restore it in a <c>finally</c> block, exactly like
    /// <c>LifecycleTools.SelectConverter</c> (another process-wide static test seam); see
    /// <c>CrossProcessLockTests</c> in <c>BcWordLayout.Tests</c>.
    /// </summary>
    internal static TimeSpan MutatingTimeout { get; set; } = TimeSpan.FromSeconds(10);

    /// <summary>
    /// Acquire timeout for every READ-ONLY tool (<see cref="ToolGuards.GuardRead"/> —
    /// <c>get_layout_info</c>/<c>list_dataset_fields</c>/<c>validate_layout</c>) — deliberately shorter than
    /// <see cref="MutatingTimeout"/> so a read waiting behind a slow concurrent preview/edit (which can itself
    /// legitimately run for up to a Word-COM conversion timeout, tens of seconds) degrades to a quick,
    /// actionable <c>file_locked</c> instead of blocking for the mutating timeout's full duration. Mutable for
    /// the same test-substitution reason as <see cref="MutatingTimeout"/>.
    /// </summary>
    internal static TimeSpan ReadTimeout { get; set; } = TimeSpan.FromSeconds(5);

    /// <summary>
    /// Attempts to acquire the cross-process lock for <paramref name="layoutPath"/> within
    /// <paramref name="timeout"/>. ALWAYS dispose the returned <see cref="Handle"/> (via <c>using</c>) — safe
    /// to do whether or not <see cref="Handle.Acquired"/> is true.
    /// </summary>
    internal static Handle TryAcquire(string layoutPath, TimeSpan timeout)
    {
        var mutex = new Mutex(initiallyOwned: false, MutexName(layoutPath));
        bool owned;
        try
        {
            owned = mutex.WaitOne(timeout);
        }
        catch (AbandonedMutexException)
        {
            // See this type's own remarks: an abandoned mutex still transfers ownership to US here, and the
            // layout file cannot be left half-committed by the dead previous holder.
            owned = true;
        }

        if (!owned)
        {
            mutex.Dispose();
            return default;
        }

        return new Handle(mutex);
    }

    /// <summary>
    /// Builds the well-known cross-process mutex name for <paramref name="layoutPath"/> — see this type's own
    /// remarks for the naming/scope/case-folding rationale. Internal (not private) so it is directly unit-
    /// testable (mirrors <c>LifecycleTools.PreviewPathHash</c>'s own reason for being internal).
    /// </summary>
    internal static string MutexName(string layoutPath)
    {
        var normalized = Path.GetFullPath(layoutPath).ToUpperInvariant();
        var hash = SHA256.HashData(Encoding.UTF8.GetBytes(normalized));
        return "Local\\bcwl-edit-" + Convert.ToHexString(hash)[..24].ToLowerInvariant();
    }

    /// <summary>
    /// Disposable handle to a (possibly not-)acquired cross-process lock. <see cref="Dispose"/> is always
    /// safe to call: <see cref="TryAcquire"/>'s failed-acquire path already disposes the underlying
    /// <see cref="Mutex"/> itself and returns <c>default</c>, whose <c>_mutex</c> field is null, so disposing
    /// a non-acquired handle is a no-op rather than a double-dispose.
    /// </summary>
    internal readonly struct Handle : IDisposable
    {
        private readonly Mutex? _mutex;

        internal Handle(Mutex mutex) => _mutex = mutex;

        /// <summary>True when the lock was actually acquired (a failed/<c>default</c> handle reports false).</summary>
        internal bool Acquired => _mutex is not null;

        public void Dispose()
        {
            if (_mutex is null)
            {
                return;
            }

            try
            {
                _mutex.ReleaseMutex();
            }
            finally
            {
                _mutex.Dispose();
            }
        }
    }
}
