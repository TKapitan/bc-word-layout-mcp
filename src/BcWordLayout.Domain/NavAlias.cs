namespace BcWordLayout.Domain;

/// <summary>
/// The <c>#Nav:</c> convention BC's own Word add-in stamps onto every control it creates, and the one
/// <see cref="SdtFactory"/> reproduces: a control's <c>w:alias</c> is <c>#Nav: &lt;dataset path&gt;</c>
/// (e.g. <c>#Nav: /Header/CustomerAddress1</c>) and its <c>w:tag</c> is
/// <c>#Nav: &lt;report name&gt;/&lt;report id&gt;</c>.
/// </summary>
/// <remarks>
/// The alias matters beyond documentation: it is the ONE part of a control that survives everything. Word's
/// compatibility downgrade strips a repeater's <c>w15:repeatingSection</c> and even its
/// <c>w15:dataBinding</c>, but leaves <c>w:alias</c>/<c>w:tag</c>/<c>w:id</c> intact — which is what lets
/// <see cref="LayoutValidator"/> recognise the wreckage as a former repeater, and what lets
/// <see cref="LayoutEditor"/> find a parent repeater's bound data item when inserting a nested row.
/// </remarks>
public static class NavAlias
{
    /// <summary>The prefix both the alias and the tag carry.</summary>
    public const string Prefix = "#Nav:";

    /// <summary>
    /// The dataset path <paramref name="alias"/> names, or <c>null</c> when it is null or does not follow the
    /// convention. Whitespace after the prefix is trimmed (the corpus writes one space; nothing should depend
    /// on that being exactly one).
    /// </summary>
    public static string? DatasetPath(string? alias) =>
        alias is not null && alias.StartsWith(Prefix, StringComparison.Ordinal)
            ? alias[Prefix.Length..].Trim()
            : null;
}
