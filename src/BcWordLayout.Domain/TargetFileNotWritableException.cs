namespace BcWordLayout.Domain;

/// <summary>
/// Thrown when the operating system refuses, with access denied, the write that would replace the
/// caller's target file: the atomic rename every writer in this codebase commits with (an edit's stage onto
/// <c>layoutPath</c>, <see cref="LayoutBuilder.Create"/>'s build onto <c>outputPath</c>,
/// <c>MergeEngine.Merge</c>'s stage onto its output), or, for an edit, the stage copy into the target's own
/// folder. <see cref="FilePath"/> names that target; the target is unchanged when this is thrown.
/// </summary>
/// <remarks>
/// <para>
/// WHY A DEDICATED TYPE. On Windows a rename onto a file fails with <see cref="UnauthorizedAccessException"/>,
/// not a sharing-violation <see cref="IOException"/>, in each of three everyday situations: the file is open
/// in Word (which holds it for read/write while sharing only read), it carries the read-only attribute (a
/// TFVC/Perforce workspace file that is not checked out), or the account lacks write permission there.
/// <c>ToolGuards.Guard</c> mapped only the sharing-violation shape to <c>file_locked</c>, so all three
/// surfaced as <c>internal_error</c> with a "likely a bug in the tool" hint (GitHub issue #94).
/// <c>Guard</c> maps THIS type to <c>file_locked</c> with a hint naming the three causes.
/// </para>
/// <para>
/// SCOPED TO THE TARGET, NOT A GLOBAL MAPPING. <see cref="UnauthorizedAccessException"/> itself stays unmapped
/// (it reports <c>internal_error</c>): the same type is thrown when a SOURCE file cannot be read, where "close
/// it in Word" would be the wrong advice. Only a throw site that knows the denied operation writes the
/// caller's target wraps it in this type.
/// </para>
/// </remarks>
public sealed class TargetFileNotWritableException : Exception
{
    public TargetFileNotWritableException(string filePath, UnauthorizedAccessException innerException)
        : base(
            $"Could not write '{filePath}': {innerException?.Message} The file was left unchanged.",
            innerException)
    {
        FilePath = filePath;
    }

    /// <summary>Full path of the target file the call could not replace.</summary>
    public string FilePath { get; }
}
