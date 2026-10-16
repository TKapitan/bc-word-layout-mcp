namespace BcWordLayout.Domain;

/// <summary>
/// Thrown by <see cref="LayoutBuilder.Create"/> when the layout it has just assembled fails its own
/// <c>OpenXmlValidator</c> self-check, which means an internal bug, never a problem with the caller's input
/// (<see cref="SdtFactory"/> and <see cref="LayoutBuilder"/> are engineered to emit only valid OOXML).
/// Nothing is written when it is thrown.
/// </summary>
/// <remarks>
/// WHY A DEDICATED TYPE (GitHub issue #95). This gate used to throw <see cref="FileFormatException"/>, the
/// same type System.IO.Packaging throws for a caller file that is not a <c>.docx</c> at all. Reporting
/// those caller files correctly (as <c>invalid_layout</c>) must never relabel this bug as the caller's
/// mistake. <c>ToolGuards.Guard</c> deliberately has no branch for this type, so it reports
/// <c>internal_error</c>, the code whose hint asks for a bug report.
/// </remarks>
public sealed class BuildSelfCheckFailedException : Exception
{
    public BuildSelfCheckFailedException(string message)
        : base(message)
    {
    }
}
