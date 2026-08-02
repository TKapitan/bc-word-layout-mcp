namespace BcWordLayout.Tests;

/// <summary>
/// A <see cref="FactAttribute"/> that runs only on Windows — the test is reported as skipped (with
/// the reason below) everywhere else. For tests that assert WINDOWS-SPECIFIC filesystem semantics
/// the product deliberately does not promise on other platforms (the Linux CI leg is a rot-guard,
/// not a support claim — see ci.yml's header and D8): mandatory share-mode locking surfacing as
/// <c>file_locked</c> (on POSIX the sharing-violation HResults never occur, and Guard's documented
/// behavior is <c>internal_error</c> — see <c>ToolGuards.IsSharingOrLockViolation</c>), drive-root
/// path shapes like <c>C:\</c>, and directories made undeletable by an open handle (POSIX happily
/// deletes open files). Tests whose inputs are merely path-shaped but platform-neutral should use
/// <c>Path.Combine</c> instead of this attribute — skipping is only for semantics that cannot be
/// arranged off-Windows at all.
/// </summary>
[AttributeUsage(AttributeTargets.Method)]
public sealed class WindowsOnlyFactAttribute : FactAttribute
{
    public WindowsOnlyFactAttribute()
    {
        if (!OperatingSystem.IsWindows())
        {
            Skip = "Windows-only filesystem semantics (share-mode locking / drive roots / open-handle delete blocking) - see this attribute's remarks.";
        }
    }
}
