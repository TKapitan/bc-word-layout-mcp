namespace BcWordLayout.McpHost;

/// <summary>
/// The <c>error.code</c> values a tool can return. They are public API (semver; see the README's "Public
/// contract and versioning" and ADR-0004), so every failure envelope takes its code from here, and
/// <c>DocDriftGuardTests</c> requires the companion skill's §5 to list exactly these and no envelope to spell
/// a code out as a literal (GitHub issue #114). A new code therefore starts here and cannot ship undocumented.
/// </summary>
internal static class ErrorCodes
{
    /// <summary>A file the call names does not exist.</summary>
    internal const string FileNotFound = "file_not_found";

    /// <summary>A layout, template or schema file the call passed is not usable as one.</summary>
    internal const string InvalidLayout = "invalid_layout";

    /// <summary>An argument is missing, malformed, out of range, or names something the layout cannot take.</summary>
    internal const string InvalidArgument = "invalid_argument";

    /// <summary>A control, table, part or dataset path the call names does not exist in the layout.</summary>
    internal const string NotFound = "not_found";

    /// <summary>The edit would introduce a new structural error; the file is untouched.</summary>
    internal const string EditWouldCorrupt = "edit_would_corrupt";

    /// <summary>The edit would leave a table row out of step with its grid; the file is untouched.</summary>
    internal const string EditWouldBreakTable = "edit_would_break_table";

    /// <summary>A file the call reads or must replace is busy or not writable.</summary>
    internal const string FileLocked = "file_locked";

    /// <summary><c>create_layout</c>: the template carries bound controls that would go stale.</summary>
    internal const string TemplateNotUnbound = "template_not_unbound";

    /// <summary><c>render_preview_pages</c>: the PDF could not be rendered (or the page range is past its end).</summary>
    internal const string RasterizeFailed = "rasterize_failed";

    /// <summary>A bug in the server, not in the call.</summary>
    internal const string InternalError = "internal_error";
}
