namespace BcWordLayout.Domain;

/// <summary>
/// Thrown when an input trips one of the ceilings in <see cref="ResourceLimits"/> — a custom XML part or
/// schema file bigger than <see cref="ResourceLimits.MaxCustomXmlPartBytes"/>, a package with more than
/// <see cref="ResourceLimits.MaxCustomXmlParts"/> custom XML parts, or a schema/document/part-graph nested
/// deeper than <see cref="ResourceLimits.MaxSchemaDepth"/>/<see cref="ResourceLimits.MaxElementNestingDepth"/>/
/// <see cref="ResourceLimits.MaxPartGraphDepth"/> — see <see cref="ResourceLimits"/> for the two crash
/// classes those caps exist to convert into ordinary, catchable failures.
/// </summary>
/// <remarks>
/// <para>
/// Derives directly from <see cref="Exception"/> (NOT <see cref="InvalidDataException"/> — that BCL type is
/// <c>sealed</c>, so no custom exception can derive from it) — the same shape <see cref="NotFoundException"/>
/// already uses for exactly the same reason: a typed marker a catch site can SWITCH ON, rather than a
/// message-text pattern. The one place that matters is <c>ToolGuards.Guard</c>, which catches THIS type
/// explicitly, ahead of its generic <c>catch (InvalidDataException)</c> branch (every OTHER malformed-layout
/// throw site still goes through that one unchanged).
/// </para>
/// <para>
/// WHY A DEDICATED TYPE (mirrors <see cref="NotFoundException"/>'s own
/// rationale). <c>ToolGuards.Guard</c>'s plain <c>invalid_layout</c> hint talks about a missing dataset part
/// or wrong namespace — accurate for most <see cref="InvalidDataException"/>s, but MISLEADING for a
/// size/part-count/depth rejection, whose real fix is "the file is too big or too deeply nested, not
/// structurally wrong". <c>Guard</c> gives THIS TYPE its own tailored hint — by switching on the TYPE, never
/// by pattern-matching <see cref="Exception.Message"/> substrings (the exact string-coupling
/// already eliminated for <c>not_found</c>/<c>invalid_argument</c>).
/// </para>
/// </remarks>
public sealed class ResourceLimitExceededException : Exception
{
    public ResourceLimitExceededException(string message)
        : base(message)
    {
    }
}
