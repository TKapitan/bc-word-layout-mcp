using System.Xml;

namespace BcWordLayout.Domain;

/// <summary>
/// Discriminates WHY an <see cref="InvalidInputFileException"/>'s file could not be read, so
/// <c>BcWordLayout.McpHost.Tools.ToolGuards.Guard</c> can pick the matching hint by switching on this value
/// rather than on the exception's free-text message (the <see cref="NotFoundTarget"/> pattern).
/// </summary>
public enum InputFileProblem
{
    /// <summary>
    /// Opening the file as a Word package failed before any part could be read: it is not a zip container
    /// at all (a schema <c>.xml</c> or a text file given a <c>.docx</c> name), an empty file, a truncated copy,
    /// or a package too damaged to open.
    /// </summary>
    NotAWordPackage,

    /// <summary>
    /// A standalone XML file (a schema <c>.xml</c>) is not well-formed XML: truncated, an encoding declaration
    /// that does not match its bytes, or a DTD, which every loader here refuses.
    /// </summary>
    MalformedXmlFile,

    /// <summary>
    /// The file opened as a Word package, but one of its XML parts is not well-formed (a damaged package).
    /// The OpenXML SDK parses parts lazily, so this surfaces at the first read of the broken part.
    /// </summary>
    MalformedPackagePart,
}

/// <summary>
/// Thrown when a caller-supplied layout, template or schema file is not the kind of file its parameter takes
/// (see <see cref="Problem"/>). <see cref="FilePath"/> names the caller's own file, even when the failing
/// read was of a private copy of it (an edit working copy, a merge stage, a template's build file).
/// <c>ToolGuards.Guard</c> maps it to <c>invalid_layout</c> with a hint naming that file and, for XML, the
/// line and position.
/// </summary>
/// <remarks>
/// <para>
/// WHY A DEDICATED TYPE (GitHub issue #95). The BCL types these failures arrive as,
/// <see cref="FileFormatException"/> from System.IO.Packaging and <see cref="XmlException"/> from the XML
/// reader, had no branch in <c>Guard</c>, so a caller's bad input was reported as <c>internal_error</c> with a
/// hint calling it "likely a bug in the tool". Mapping either BCL type wholesale would be wrong:
/// <see cref="LayoutBuilder"/>'s own internal-bug gate used to throw <see cref="FileFormatException"/>, and
/// neither type says WHICH file was bad. So the throw sites that know the file, chiefly
/// <see cref="LayoutPackage.Open"/> and <see cref="SchemaProvider.FromSchemaXml"/>, wrap the BCL exception in
/// this type, which carries it.
/// </para>
/// <para>
/// NOT FOR AN AUXILIARY PARAMETER FILE. A malformed <c>dataOverridesPath</c> file is reported as an
/// <see cref="ArgumentException"/> for that parameter instead (<c>invalid_argument</c>), because the fix is
/// "pass a different data file", not "repair the layout".
/// </para>
/// </remarks>
public sealed class InvalidInputFileException : Exception
{
    public InvalidInputFileException(string filePath, InputFileProblem problem, string message, Exception? innerException = null)
        : base(message, innerException)
    {
        FilePath = filePath;
        Problem = problem;
        if (innerException is XmlException { LineNumber: > 0 } xml)
        {
            LineNumber = xml.LineNumber;
            LinePosition = xml.LinePosition;
        }
    }

    /// <summary>The caller-supplied path of the file that could not be read.</summary>
    public string FilePath { get; }

    /// <summary>Why the file could not be read.</summary>
    public InputFileProblem Problem { get; }

    /// <summary>1-based line of an XML error, when the reader reported one.</summary>
    public int? LineNumber { get; }

    /// <summary>1-based position within <see cref="LineNumber"/> of an XML error, when the reader reported one.</summary>
    public int? LinePosition { get; }
}
