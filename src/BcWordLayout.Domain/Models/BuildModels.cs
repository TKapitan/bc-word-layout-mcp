namespace BcWordLayout.Domain.Models;

/// <summary>
/// Describes a newly created BC Word layout (see <see cref="BcWordLayout.Domain.LayoutBuilder.Create"/>):
/// enough detail for a caller (typically the <c>create_layout</c> MCP tool) to report exactly what was
/// produced without re-reading the file.
/// </summary>
public sealed class CreateResult
{
    /// <summary>Absolute path to the created <c>.docx</c> layout.</summary>
    public required string OutputPath { get; init; }

    /// <summary>The report's name, parsed from the dataset namespace (see <see cref="ReportIdentity"/>).</summary>
    public required string ReportName { get; init; }

    /// <summary>The report's id, parsed from the dataset namespace.</summary>
    public required string ReportId { get; init; }

    /// <summary>The full dataset namespace URI carried by the attached BC custom XML part.</summary>
    public required string Namespace { get; init; }

    /// <summary>
    /// The freshly generated <c>ds:itemID</c> GUID (<c>{GUID}</c>, uppercase) of the created BC custom XML
    /// part — what any control subsequently inserted by <see cref="BcWordLayout.Domain.LayoutEditor"/> binds
    /// its <c>storeItemID</c> to.
    /// </summary>
    public required string StoreItemId { get; init; }

    /// <summary>True when a <c>templatePath</c> was supplied (the output started life as a copy of it).</summary>
    public required bool UsedTemplate { get; init; }

    /// <summary>
    /// True when the template already had its own BC dataset custom XML part, which was removed and
    /// replaced by the one built from <c>schemaSource</c>. Always false when <see cref="UsedTemplate"/> is
    /// false (a freshly created document never has a pre-existing BC part to replace).
    /// </summary>
    public required bool ReplacedExistingBcPart { get; init; }

    /// <summary>
    /// The result of running <see cref="BcWordLayout.Domain.LayoutValidator.Quick"/> against the built
    /// layout while the package was still open (before it was committed to <see cref="OutputPath"/>). A
    /// freshly created (non-template) layout always passes; the one case that used to reach this property
    /// with errors — a <c>templatePath</c> whose own bound content controls go stale once its BC part is
    /// REPLACED — is now refused outright before <see cref="LayoutBuilder.Create"/> ever returns (see
    /// <see cref="BcWordLayout.Domain.TemplateNotUnboundException"/>). Still worth checking rather than
    /// assuming a successful call is fully clean: a template's pre-existing content can carry its own
    /// WARNING-level finding (e.g. <c>attached-template</c>), and a template whose bound controls were never
    /// matched to any BC part in the first place — no part existed for <see cref="LayoutBuilder"/> to
    /// replace, so <see cref="ReplacedExistingBcPart"/> is false and the refusal above does not key on it —
    /// is not covered by that refusal either.
    /// </summary>
    public required ValidationResult QuickValidation { get; init; }
}
