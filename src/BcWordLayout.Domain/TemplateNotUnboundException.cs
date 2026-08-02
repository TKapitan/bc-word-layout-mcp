namespace BcWordLayout.Domain;

/// <summary>
/// Thrown by <see cref="LayoutBuilder.Create"/> when <c>templatePath</c> turns out to be a FULL BC layout —
/// not the unbound branded/styled shell the tool expects — in a way that would leave real damage behind:
/// it already carried its own BC dataset custom XML part, AND that part's replacement (which
/// <see cref="LayoutBuilder.Create"/> always performs when attaching <c>schemaSource</c>) leaves at least one
/// of the template's own pre-existing bound content controls stale — each one's <c>storeItemID</c> pointed at
/// the now-removed part, and/or its XPath no longer resolves against <c>schemaSource</c>'s shape.
/// </summary>
/// <remarks>
/// <para>
/// WHY REFUSE INSTEAD OF REPORTING A WARNING. Before this type existed,
/// <see cref="LayoutBuilder.Create"/> built the layout anyway, replaced the part, and surfaced the damage
/// only as data — a non-zero <c>QuickValidation.ErrorCount</c> plus a <c>StaleControlsWarning</c> string — on
/// an otherwise <c>Ok=true</c> result. A caller that (reasonably) treats <c>Ok=true</c> as "it worked" ships a
/// layout with dozens of dangling bindings. There is also no way to FIX this in place: every stale control's
/// <c>storeItemID</c> names a part this call just deleted, and BC re-binds by XPath against whatever schema
/// the part actually describes, not by any recoverable mapping back to the template's original one — so
/// silently rebinding or dropping the controls here would be guessing at the caller's intent, not fixing a
/// bug. Refusing instead is the "refuse rather than guess" principle this codebase already applies elsewhere
/// (see e.g. <c>LocationResolver</c>'s <c>AfterControlPosition</c> rejection): <see cref="LayoutBuilder.Create"/>
/// throws BEFORE the temp-file build is ever moved to <c>outputPath</c> (see <c>Create</c>'s own atomic-write
/// remarks), so the caller's output path is left completely untouched, exactly like every other Create-time
/// failure.
/// </para>
/// <para>
/// NOT THROWN merely because the template HAD a BC part (see <see cref="CreateResult.ReplacedExistingBcPart"/>,
/// which still fires for this case) — only when replacing it actually leaves something broken. A template
/// whose BC part had zero bound controls of its own (nothing left stale) is not this failure at all;
/// <see cref="LayoutBuilder.Create"/> keys the refusal on the built output's OWN post-build
/// <see cref="LayoutValidator.Quick"/> pass reporting at least one error, not on the part's mere presence.
/// (That zero-controls case is the ONLY replace-and-succeed shape: the fresh part always gets a newly minted
/// <c>storeItemID</c>, so ANY surviving bound control fails the store-item-id check no matter how similar the
/// schemas are — "same schema, controls still valid" cannot happen.)
/// </para>
/// </remarks>
public sealed class TemplateNotUnboundException : Exception
{
    public TemplateNotUnboundException(string message, int errorCount)
        : base(message)
    {
        ErrorCount = errorCount;
    }

    /// <summary>
    /// The stale-control quick-validation error count the offending build produced — already folded into
    /// <see cref="Exception.Message"/>, but exposed here too for a caller that wants it structurally rather
    /// than parsed out of free text.
    /// </summary>
    public int ErrorCount { get; }
}
