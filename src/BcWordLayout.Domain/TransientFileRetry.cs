namespace BcWordLayout.Domain;

/// <summary>
/// Runs a file copy/move/delete, absorbing the TRANSIENT access-denied / sharing-violation errors an
/// antivirus or indexer scan holds on a freshly written <c>.docx</c> for a few tens of milliseconds
/// (observed as one-off "Access to the path is denied." failures on Windows — Defender scans every new
/// Office file, and a path that was written moments ago by a previous tool call is exactly what it is
/// busy scanning). Bounded: a few short backoff attempts, then the original exception propagates
/// unchanged, so a PERSISTENT hold (the file genuinely open in Word) still surfaces as the same
/// file-locked/IO failure it always did. Shared by every commit path that writes a layout —
/// <c>BcWordLayout.McpHost.Tools.ToolGuards</c>'s working-copy/stage/rename steps and
/// <see cref="LayoutBuilder"/>'s temp-build/atomic-move steps.
/// </summary>
public static class TransientFileRetry
{
    private const int MaxAttempts = 4;
    private const int ErrorSharingViolation = unchecked((int)0x80070020);

    /// <summary>Runs <paramref name="fileOperation"/>, retrying transient denials with short backoff.</summary>
    public static void Run(Action fileOperation)
    {
        ArgumentNullException.ThrowIfNull(fileOperation);

        for (var attempt = 1; ; attempt++)
        {
            try
            {
                fileOperation();
                return;
            }
            catch (Exception ex) when (attempt < MaxAttempts
                && (ex is UnauthorizedAccessException || (ex is IOException && ex.HResult == ErrorSharingViolation)))
            {
                Thread.Sleep(50 * attempt);
            }
        }
    }
}
