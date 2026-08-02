using DocumentFormat.OpenXml;
using BcWordLayout.Domain.Models;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Domain;

/// <summary>
/// THE single source of truth for classifying a <c>w:sdt</c> from its <c>w:sdtPr</c> markers, and for
/// reading the handful of common sdtPr properties (tag/alias/id/dataBinding) that every reader, editor,
/// validator, and the merge engine each need. Consolidates four independently hand-synced (and already
/// slightly diverged) <c>Classify</c>/<c>ClassifyKind</c> copies — <c>LayoutReader</c>, <c>LayoutEditor</c>,
/// <c>TableStructureReader</c>, and <c>BcWordLayout.Merge.MergeEngine</c> — plus the
/// <c>FindBinding</c>/<c>ReadControlId</c>/<c>FirstChild</c>/<c>HasChild</c>/<c>Attr</c> helper family that
/// had been copy-pasted (not just the classification ladder) into <c>LayoutRefresher</c>,
/// <c>LayoutValidator</c>, <c>LocationResolver</c>, <c>TableGridNavigator</c>, and
/// <c>PlainTextNestingGuard</c> too.
/// </summary>
/// <remarks>
/// <para><b>Divergences found while consolidating, and how each was resolved:</b></para>
/// <list type="number">
/// <item>
/// <b>Drift (fixed):</b> <c>TableStructureReader.ClassifyKind</c> had no <c>repeatingSectionItem</c> branch
/// — an item wrapper that happened to carry its own <c>w:dataBinding</c> would have fallen through to
/// Field/Label there, while the other three copies all treated it as a non-surfaced structural wrapper
/// (mapped to <see cref="ControlKind.Unbound"/>). <see cref="ClassifyControlKind"/> now applies the SAME
/// branch order everywhere, so this can no longer happen regardless of which caller classifies the sdt.
/// Latent in the real corpus checked so far — a <c>repeatingSectionItem</c> sdt never itself carries a
/// binding — but no longer a trap for a future corpus/hand-authored layout that does.
/// </item>
/// <item>
/// <b>Cosmetic (unified):</b> <c>LayoutEditor.ClassifyKind</c> returned a <c>string</c>
/// (<c>ControlKind.X.ToString()</c>) rather than the <see cref="ControlKind"/> enum the other three copies
/// used. <see cref="ClassifyControlKind"/> returns <see cref="ControlKind"/>; callers that need the string
/// form (e.g. <c>EditResult.Kind</c>) call <c>.ToString()</c> at their own boundary, exactly as before.
/// </item>
/// <item>
/// <b>Deliberate, kept explicit:</b> a repeating section's own row xpath (read by
/// <c>MergeEngine.ProcessRepeater</c> and <c>LayoutValidator</c>'s repeater-shape diagnostics) is looked up
/// via <see cref="FindRepeaterBinding"/> — <c>w15:dataBinding</c> ONLY, never falling back to a legacy
/// <c>w:dataBinding</c> — deliberately narrower than <see cref="FindBinding(SdtProperties)"/>'s
/// field/label-oriented "prefer legacy, else w15" search. See <see cref="FindRepeaterBinding"/>'s own
/// remarks for why. This is not a drift fix: both call sites already special-cased w15-only before this
/// consolidation; it is now one named method instead of two hand-written inline lookups.
/// </item>
/// </list>
/// </remarks>
internal static class SdtInspector
{
    /// <summary>
    /// The marker-driven classification ladder every consumer's control-flow branches on: what KIND of
    /// <c>w:sdt</c> is this, from its <c>w:sdtPr</c> children alone (no dataset/schema knowledge). See
    /// <see cref="ClassifyControlKind"/> for the higher-level mapping onto the public
    /// <see cref="ControlKind"/> model (which additionally applies <see cref="LabelConvention"/> to split
    /// <see cref="Classification.Bound"/> into Field vs. Label).
    /// </summary>
    internal enum Classification
    {
        /// <summary><c>w15:repeatingSection</c> — the repeating-section control itself.</summary>
        Repeater,

        /// <summary>
        /// <c>w15:repeatingSectionItem</c> — the row TEMPLATE a repeater wraps; a structural wrapper only,
        /// never surfaced as its own inventory entry by any caller (its content flows through transparently).
        /// </summary>
        RepeaterItem,

        /// <summary><c>w:picture</c> — checked before <see cref="Bound"/> because a picture control also carries a <c>w:dataBinding</c>.</summary>
        Picture,

        /// <summary>Has a <c>w:dataBinding</c> or <c>w15:dataBinding</c> (see <see cref="FindBinding(SdtProperties)"/>) and none of the markers above.</summary>
        Bound,

        /// <summary>No recognized marker and no binding — e.g. a page-number field, a locked <c>w:group</c> wrapper.</summary>
        Unbound,
    }

    /// <summary>
    /// Classifies an sdt from its <paramref name="pr"/> (null when the sdt has no <c>w:sdtPr</c> at all,
    /// which classifies as <see cref="Classification.Unbound"/>). Branch order matters and is load-bearing:
    /// <see cref="Classification.Repeater"/>/<see cref="Classification.RepeaterItem"/> are checked before
    /// <see cref="Classification.Picture"/>, which is itself checked before <see cref="Classification.Bound"/>
    /// — a picture control also carries a <c>w:dataBinding</c>, so the picture marker must win, and (per the
    /// verified corpus) a repeating section's own sdtPr also carries a <c>w15:dataBinding</c> alongside its
    /// <c>w15:repeatingSection</c> marker, so that must be checked first of all.
    /// </summary>
    internal static Classification Classify(SdtProperties? pr)
    {
        if (pr is null)
        {
            return Classification.Unbound;
        }

        if (HasChild(pr, "repeatingSection", OoxmlNames.W15))
        {
            return Classification.Repeater;
        }

        if (HasChild(pr, "repeatingSectionItem", OoxmlNames.W15))
        {
            return Classification.RepeaterItem;
        }

        if (HasChild(pr, "picture", OoxmlNames.W))
        {
            return Classification.Picture;
        }

        if (FindBinding(pr) is not null)
        {
            return Classification.Bound;
        }

        return Classification.Unbound;
    }

    /// <summary>
    /// Maps <see cref="Classify"/> onto the public <see cref="ControlKind"/> model: a bound control is
    /// further split into <see cref="ControlKind.Label"/>/<see cref="ControlKind.Field"/> by the active
    /// <see cref="LabelConvention.Current"/> against its binding's full xpath segments (not just the leaf
    /// name — a custom convention's <see cref="LabelConvention.LabelsDataItemName"/> rule also needs the
    /// immediate parent segment), and <see cref="Classification.RepeaterItem"/> — a structural wrapper no
    /// caller surfaces as its own control — maps to <see cref="ControlKind.Unbound"/> (matching every
    /// existing caller's treatment of it; see this type's own remarks, divergence 1).
    /// </summary>
    internal static ControlKind ClassifyControlKind(SdtProperties? pr) => Classify(pr) switch
    {
        Classification.Repeater => ControlKind.Repeater,
        Classification.Picture => ControlKind.Picture,
        Classification.Bound => LabelConvention.Current.IsLabelPath(
            BindingXPath.Segments(Attr(FindBinding(pr!)!, "xpath", OoxmlNames.W)))
            ? ControlKind.Label
            : ControlKind.Field,
        _ => ControlKind.Unbound,
    };

    /// <summary>Convenience overload of <see cref="ClassifyControlKind(SdtProperties?)"/> starting from the sdt itself.</summary>
    internal static ControlKind ClassifyControlKind(SdtElement sdt) => ClassifyControlKind(sdt.GetFirstChild<SdtProperties>());

    /// <summary>True when <paramref name="sdt"/> itself carries the <c>w15:repeatingSection</c> marker.</summary>
    internal static bool IsRepeater(SdtElement sdt) => HasPrChild(sdt, "repeatingSection", OoxmlNames.W15);

    /// <summary>True when <paramref name="sdt"/> itself carries the <c>w15:repeatingSectionItem</c> marker.</summary>
    internal static bool IsRepeaterItem(SdtElement sdt) => HasPrChild(sdt, "repeatingSectionItem", OoxmlNames.W15);

    /// <summary>The nearest ancestor sdt (if any) for which <see cref="IsRepeater"/> is true.</summary>
    internal static SdtElement? NearestRepeaterAncestor(SdtElement sdt) =>
        sdt.Ancestors<SdtElement>().FirstOrDefault(IsRepeater);

    /// <summary>
    /// Returns the binding element for a FIELD/LABEL/PICTURE control: <c>w:dataBinding</c> preferred, else
    /// <c>w15:dataBinding</c> (real corpus field/label controls use the legacy element; some — none seen so
    /// far, but the fallback costs nothing — could plausibly use w15 instead). NOT the right lookup for a
    /// repeating section's own row binding — see <see cref="FindRepeaterBinding"/>.
    /// </summary>
    internal static OpenXmlElement? FindBinding(SdtProperties pr) =>
        FirstChild(pr, "dataBinding", OoxmlNames.W) ?? FirstChild(pr, "dataBinding", OoxmlNames.W15);

    /// <summary>Convenience overload of <see cref="FindBinding(SdtProperties)"/> starting from the sdt itself.</summary>
    internal static OpenXmlElement? FindBinding(SdtElement sdt) =>
        sdt.GetFirstChild<SdtProperties>() is { } pr ? FindBinding(pr) : null;

    /// <summary>
    /// Returns a repeating section's (or, equivalently, a repeater's item template's) OWN row/data-item
    /// binding element — <c>w15:dataBinding</c> ONLY, deliberately never falling back to a legacy
    /// <c>w:dataBinding</c> the way <see cref="FindBinding(SdtProperties)"/> does for field/label controls.
    /// This is DELIBERATE, not an oversight: the verified corpus shows a repeating section's own binding is
    /// always <c>w15:dataBinding</c> with an UNINDEXED final xpath step (see the
    /// <c>bc-word-layout-ooxml-facts</c> project memory), never a legacy one. If a hand-authored or future
    /// layout ever put a stray legacy <c>w:dataBinding</c> on a repeater's own <c>w:sdtPr</c>, silently
    /// falling back to it here would hand <c>MergeEngine.ProcessRepeater</c>/<c>XPathReanchor</c> the WRONG
    /// xpath for the row context every inner binding in that row is re-anchored against — a much worse
    /// failure than "repeater xpath missing", which the w15-only lookup degrades to instead (the repeater
    /// is then skipped, exactly as a repeater with no binding at all already is). Used by
    /// <c>MergeEngine.ProcessRepeater</c> (the row xpath driving row-cloning) and
    /// <c>LayoutValidator.RepeaterXPath</c> (a diagnostic location string only).
    /// </summary>
    internal static OpenXmlElement? FindRepeaterBinding(SdtProperties pr) =>
        FirstChild(pr, "dataBinding", OoxmlNames.W15);

    /// <summary>The sdt's <c>w:sdtPr/w:alias/@w:val</c>, or null.</summary>
    internal static string? ReadAlias(SdtElement sdt) =>
        sdt.GetFirstChild<SdtProperties>()?.GetFirstChild<SdtAlias>()?.Val?.Value;

    /// <summary>The sdt's <c>w:sdtPr/w:tag/@w:val</c>, or null.</summary>
    internal static string? ReadTag(SdtElement sdt) =>
        sdt.GetFirstChild<SdtProperties>()?.GetFirstChild<Tag>()?.Val?.Value;

    /// <summary>The sdt's <c>w:sdtPr/w:id/@w:val</c>, or null when absent.</summary>
    internal static int? ReadControlId(SdtElement sdt) =>
        sdt.GetFirstChild<SdtProperties>()?.GetFirstChild<SdtId>()?.Val?.Value;

    /// <summary>The sdt's binding xpath (see <see cref="FindBinding(SdtElement)"/>), or null when unbound.</summary>
    internal static string? ReadXPath(SdtElement sdt)
    {
        var binding = FindBinding(sdt);
        return binding is null ? null : Attr(binding, "xpath", OoxmlNames.W);
    }

    /// <summary>The sdt's binding <c>storeItemID</c> (see <see cref="FindBinding(SdtElement)"/>), or null when unbound.</summary>
    internal static string? ReadStoreItemId(SdtElement sdt)
    {
        var binding = FindBinding(sdt);
        return binding is null ? null : Attr(binding, "storeItemID", OoxmlNames.W);
    }

    /// <summary>
    /// The BC dataset namespace URI declared by the sdt's binding <c>w:prefixMappings</c>, or null when the
    /// control is unbound or its prefixMappings names no BC namespace.
    /// </summary>
    /// <remarks>
    /// The attribute's normal form is <c>xmlns:ns0='&lt;uri&gt;'</c>, but real base-app layouts also ship it
    /// as a BARE uri with no <c>xmlns:</c> declaration at all while the XPath still uses the <c>ns0:</c>
    /// prefix (`StandardSalesInvoiceVatSpec.docx`'s Mini_Sales_Invoice bindings) — so the URI is extracted by
    /// pattern rather than by parsing an xmlns declaration. A single prefixMappings may declare several
    /// prefixes; only the BC one is of interest here, and only the first is returned (no real layout declares
    /// two different BC namespaces in one binding).
    /// </remarks>
    internal static string? ReadBindingNamespace(SdtElement sdt)
    {
        var binding = FindBinding(sdt);
        var mappings = binding is null ? null : Attr(binding, "prefixMappings", OoxmlNames.W);
        return ExtractBcNamespace(mappings);
    }

    /// <summary>
    /// Pulls the first <c>urn:microsoft-dynamics…</c> URI out of a <c>w:prefixMappings</c> value, whatever
    /// form it takes (quoted xmlns declaration or bare URI). Internal for direct testing.
    /// </summary>
    internal static string? ExtractBcNamespace(string? prefixMappings)
    {
        if (string.IsNullOrWhiteSpace(prefixMappings))
        {
            return null;
        }

        var start = prefixMappings.IndexOf(OoxmlNames.BcNamespacePrefix, StringComparison.Ordinal);
        if (start < 0)
        {
            return null;
        }

        // The URI runs to the first delimiter that cannot be part of it: the closing quote of an xmlns
        // declaration, or whitespace when it is bare. A BC namespace itself never contains either.
        var end = start;
        while (end < prefixMappings.Length
               && prefixMappings[end] is not ('\'' or '"')
               && !char.IsWhiteSpace(prefixMappings[end]))
        {
            end++;
        }

        return prefixMappings[start..end];
    }

    /// <summary>True when the sdt's binding (see <see cref="FindBinding(SdtElement)"/>) is a <c>w15:dataBinding</c> rather than legacy <c>w:dataBinding</c>.</summary>
    internal static bool UsesW15Binding(SdtElement sdt)
    {
        var binding = FindBinding(sdt);
        return binding is not null && binding.NamespaceUri == OoxmlNames.W15;
    }

    // ---- generic OOXML child/attribute lookup (namespace-aware; the OpenXml SDK's own Elements()/GetAttributes() ----
    // ---- do not offer a single-call "named child in this namespace" or "named attribute in this namespace" helper) ----

    /// <summary>True when <paramref name="el"/> has a direct child element named <paramref name="localName"/> in namespace <paramref name="ns"/>.</summary>
    internal static bool HasChild(OpenXmlElement el, string localName, string ns) =>
        FirstChild(el, localName, ns) is not null;

    /// <summary>The first direct child of <paramref name="el"/> named <paramref name="localName"/> in namespace <paramref name="ns"/>, or null.</summary>
    internal static OpenXmlElement? FirstChild(OpenXmlElement el, string localName, string ns) =>
        el.Elements().FirstOrDefault(e => e.LocalName == localName && e.NamespaceUri == ns);

    /// <summary>
    /// The <c>val</c> attribute of <paramref name="el"/>'s first child named <paramref name="localName"/>/
    /// <paramref name="ns"/>, or null. The <c>val</c> attribute itself is always read in the <c>w:</c>
    /// namespace regardless of <paramref name="ns"/> — per <see cref="OoxmlNames"/>'s own remarks, BC's
    /// attribute values live in <c>w:</c> even on a <c>w15:</c>-namespaced element.
    /// </summary>
    internal static string? ChildVal(OpenXmlElement el, string localName, string ns)
    {
        var child = FirstChild(el, localName, ns);
        return child is null ? null : Attr(child, "val", OoxmlNames.W);
    }

    /// <summary>The value of <paramref name="el"/>'s attribute named <paramref name="localName"/> in namespace <paramref name="ns"/>, or null.</summary>
    internal static string? Attr(OpenXmlElement el, string localName, string ns)
    {
        foreach (var a in el.GetAttributes())
        {
            if (a.LocalName == localName && a.NamespaceUri == ns)
            {
                return a.Value;
            }
        }

        return null;
    }

    private static bool HasPrChild(SdtElement sdt, string localName, string ns)
    {
        var pr = sdt.GetFirstChild<SdtProperties>();
        return pr is not null && HasChild(pr, localName, ns);
    }
}
