namespace BcWordLayout.Domain;

/// <summary>
/// Discriminates WHICH defect of the layout package a <see cref="LayoutDefectException"/> reports, so
/// <c>BcWordLayout.McpHost.Tools.ToolGuards.Guard</c> can pick the repair hint by switching on this value
/// rather than on the exception's free-text <see cref="Exception.Message"/> (ADR-0004) — the same pattern
/// <see cref="NotFoundTarget"/> follows for <see cref="NotFoundException"/>.
/// </summary>
public enum LayoutDefect
{
    /// <summary>
    /// The package's BC dataset custom XML part is recognisably there (its itemProps <c>ds:schemaRef</c> or
    /// its raw bytes name the BC report namespace) but is not well-formed XML, so no dataset can be read
    /// from it — typically an encoding declaration that contradicts the part's bytes after a text-editor
    /// re-save, an unescaped <c>&amp;</c>, a truncated file, or a DOCTYPE (GitHub issue #96).
    /// </summary>
    UnreadableDatasetPart,

    /// <summary>
    /// The package carries more than one BC dataset custom XML part, so an operation that copies, replaces
    /// or binds to "the" dataset part would have to guess which one is meant (GitHub issue #101). Reading
    /// still works against the part the bindings use; writing refuses (ADR-0003).
    /// </summary>
    SeveralDatasetParts,

    /// <summary>
    /// More than one content control in the package carries the <c>w:sdtPr/w:id</c> an id-addressed tool was
    /// given, so the id does not say which control is meant (GitHub issue #101).
    /// </summary>
    DuplicateControlId,
}

/// <summary>
/// Thrown when the layout PACKAGE itself carries a defect that the requested operation cannot work around —
/// as opposed to a bad argument (<see cref="ArgumentException"/>, <c>invalid_argument</c>) or a lookup that
/// found nothing (<see cref="NotFoundException"/>, <c>not_found</c>). <see cref="Defect"/> says which defect,
/// and <c>ToolGuards.Guard</c> maps every one of them to <c>invalid_layout</c> with a hint telling the caller
/// how to repair the package.
/// </summary>
/// <remarks>
/// WHY A DEDICATED TYPE. <see cref="InvalidDataException"/> already maps to <c>invalid_layout</c>, but with
/// one generic hint ("confirm the file is a BC layout containing a dataset part ... use get_layout_info"),
/// which is wrong for every defect listed in <see cref="LayoutDefect"/>: the dataset part IS there, and
/// get_layout_info is often the very tool that just failed. <see cref="InvalidDataException"/> is
/// <c>sealed</c>, so the defects cannot be told apart by subtype either; this type and its
/// <see cref="Defect"/> discriminator are what let the envelope name the actual repair.
/// </remarks>
public sealed class LayoutDefectException : Exception
{
    /// <summary>Constructs a package-defect failure carrying <paramref name="defect"/> for hint selection.</summary>
    public LayoutDefectException(string message, LayoutDefect defect)
        : base(message)
    {
        Defect = defect;
    }

    /// <summary>Which defect of the package this is — see <see cref="LayoutDefect"/>'s own members.</summary>
    public LayoutDefect Defect { get; }
}
