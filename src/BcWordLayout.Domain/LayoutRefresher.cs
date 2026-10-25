using BcWordLayout.Domain.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Ds = DocumentFormat.OpenXml.CustomXmlDataProperties;
using Office2013Word = DocumentFormat.OpenXml.Office2013.Word;

namespace BcWordLayout.Domain;

/// <summary>
/// Updates a layout's BC dataset custom XML part to a NEW schema (e.g. after the AL report dataset changed)
/// in place — content is swapped, the existing <c>ds:itemID</c> is preserved so every existing
/// <c>w:storeItemID</c> binding still links to the SAME part, and every existing control binding is
/// classified against the new schema: bindings whose XPath still resolves by element name are "remapped"
/// (kept, valid, untouched); bindings that no longer resolve are "orphaned" (reported, left in place — the
/// caller decides whether to rebind or delete them). This type works against an ALREADY-OPEN
/// <see cref="WordprocessingDocument"/> and does no file I/O of its own — opening, pre-save validation, and
/// saving are the caller's job (see <c>BcWordLayout.McpHost.Tools.LifecycleTools</c>'s <c>refresh_xml_part</c>
/// tool, which wraps this call with the same open/validate/save-or-reject safety every mutating tool uses).
/// </summary>
/// <remarks>
/// <para>
/// The <c>w16sdtdh:storeItemChecksum</c> Word 2021+ writes on a binding (6 of the 10 corpus layouts carry
/// them) is deliberately left alone, although swapping the part's content leaves it stale (GitHub issue #6).
/// It is a hash of the DATA PART's content, not of the binding: every control bound to the same part carries
/// the same value. Measured 2026-10-17 on JobQuote.docx through Word for Microsoft 365:
/// <list type="bullet">
/// <item>pristine, Word keeps the shipped value on save, so the corpus checksums are valid;</item>
/// <item>with the data changed under the old checksum, Word displays the NEW value and writes a recomputed
/// checksum on save;</item>
/// <item>with a binding's XPath re-pointed under its old checksum, Word displays the new target's value and
/// keeps the checksum, which is still correct because the part did not change.</item>
/// </list>
/// So a stale checksum never makes Word show stale data and heals on the next Word save. BC cannot be misled
/// either: it renders with real data, which never matches the sample data the checksum was computed over. A
/// sandbox upload of a binding-edited layout that kept its checksums byte for byte was accepted (issue #1's
/// rounds). Dropping the attribute would be a rewrite with no evidence of need.
/// </para>
/// </remarks>
public static class LayoutRefresher
{
    /// <summary>
    /// Refreshes <paramref name="doc"/>'s BC dataset custom XML part to the schema loaded from
    /// <paramref name="newSchemaSource"/>.
    /// </summary>
    /// <param name="doc">An already-open, editable layout package.</param>
    /// <param name="newSchemaSource">
    /// Either an absolute path to an existing <c>.docx</c> layout (its BC dataset custom XML part is located
    /// and its raw bytes copied byte-for-byte, exactly like <see cref="LayoutBuilder.Create"/>'s own
    /// <c>schemaSource</c> handling) or a standalone exported schema <c>.xml</c> (validated via
    /// <see cref="SchemaProvider.FromSchemaXml"/>; its raw bytes are used as-is).
    /// </param>
    /// <returns>
    /// A <see cref="RefreshResult"/> describing old vs new identity, the remap/orphan/new-field report, and
    /// a post-refresh <see cref="LayoutValidator.Quick"/> summary. See <see cref="RefreshResult"/>'s own
    /// remarks for exactly what counts as "remapped", "orphaned", and "new unbound".
    /// </returns>
    /// <exception cref="FileNotFoundException"><paramref name="newSchemaSource"/> does not exist.</exception>
    /// <exception cref="InvalidDataException">
    /// <paramref name="doc"/> has no main document part or no BC dataset custom XML part to refresh (a part
    /// that is there but cannot be parsed is REPAIRED, not refused - see <see cref="ResolveTarget"/>);
    /// <paramref name="newSchemaSource"/> is a <c>.docx</c> with no main document part or no BC dataset
    /// custom XML part; or is a schema <c>.xml</c> whose root is not <c>NavWordReportXmlPart</c>.
    /// </exception>
    /// <exception cref="LayoutDefectException">
    /// <paramref name="doc"/> or a <c>.docx</c> <paramref name="newSchemaSource"/> carries several BC dataset
    /// parts (<see cref="LayoutDefect.SeveralDatasetParts"/>), or the source's BC dataset part cannot be parsed.
    /// </exception>
    public static RefreshResult Refresh(WordprocessingDocument doc, string newSchemaSource)
    {
        ArgumentNullException.ThrowIfNull(doc);
        if (string.IsNullOrWhiteSpace(newSchemaSource))
        {
            throw new ArgumentException("New schema source path must not be empty.", nameof(newSchemaSource));
        }

        var main = doc.MainDocumentPart
            ?? throw new InvalidDataException("Layout has no main document part.");

        // ---- capture the OLD identity/part BEFORE anything is loaded or mutated ----
        var target = ResolveTarget(main);
        var bcPart = target.Part;
        var oldStoreItemId = target.OldIdentity.StoreItemId;
        var oldIdentity = target.OldIdentity;
        var oldColumnPaths = target.OldColumnPaths;

        // Capture the control inventory before any mutation. This is equally valid read before or after -
        // refresh never changes a binding's XPath element-name steps (only w:prefixMappings/w:tag, on a
        // namespace change, and neither is read here) - reading first just keeps "capture, then mutate" clean.
        var inventory = LayoutReader.Read(doc);

        // Load the NEW schema fully into memory (bytes + parsed tree) before the BC part is touched -
        // mirrors LayoutBuilder.Create reading schemaSource to completion before outputPath is ever touched.
        var (newBytes, newTree) = LoadNewSchema(newSchemaSource);
        var newIdentity = newTree.Report;

        var namespaceChanged = !string.Equals(oldIdentity.Namespace, newIdentity.Namespace, StringComparison.Ordinal);

        // ---- classify every existing binding against the NEW schema ----
        var remappedCount = 0;
        var orphaned = new List<OrphanedBinding>();
        foreach (var control in inventory.Controls)
        {
            if (control.XPath is null)
            {
                continue; // Unbound control: not a binding, nothing to classify.
            }

            if (LayoutValidator.Resolves(control.XPath, newTree, out _))
            {
                remappedCount++;
            }
            else
            {
                orphaned.Add(new OrphanedBinding { Alias = control.Alias, XPath = control.XPath, Part = control.Part, SdtId = control.SdtId });
            }
        }

        // ---- foreign-namespace bindings: captured BEFORE the re-point below rewrites them ----
        //
        // A binding naming neither the old nor the new namespace is orphaned onto a DIFFERENT report, and BC
        // rejects the whole layout at upload with one InvalidPrefixMapping per such binding
        // (sandbox-verified 2026-08-02, issue #1). The re-point below repairs them, so this list is a repair
        // report; it is built from the pre-mutation inventory because the namespace it reports is the one the
        // binding is about to stop naming.
        var repointedForeign = inventory.Controls
            .Where(c => c.XPath is not null
                        && c.BindingNamespace is not null
                        && !string.Equals(c.BindingNamespace, oldIdentity.Namespace, StringComparison.Ordinal)
                        && !string.Equals(c.BindingNamespace, newIdentity.Namespace, StringComparison.Ordinal))
            .Select(c => new RepointedBinding
            {
                Alias = c.Alias,
                XPath = c.XPath!,
                Part = c.Part,
                SdtId = c.SdtId,
                PreviousNamespace = c.BindingNamespace!,
            })
            .ToList();

        // ---- storeItemIDs naming no part: captured BEFORE the re-point below rewrites them ----
        //
        // Only BC-namespaced bindings, and only an ID that names NO custom XML part: Word and Business Central
        // re-bind those to the BC part anyway (sandbox pack item P08, GitHub issue #102), so pointing them at it
        // is not a guess. One naming another existing part is bound to that part's data and stays as it is.
        var partItemIds = LayoutValidator.CustomXmlPartItemIds(main);
        var repointedStoreItemIds = oldStoreItemId is null
            ? []
            : inventory.Controls
                .Where(c => c.XPath is not null
                            && c.BindingNamespace is not null
                            && c.StoreItemId is not null
                            && !partItemIds.Any(id => SchemaProvider.SameItemId(id, c.StoreItemId)))
                .Select(c => new RepointedStoreItemId
                {
                    Alias = c.Alias,
                    XPath = c.XPath!,
                    Part = c.Part,
                    SdtId = c.SdtId,
                    PreviousStoreItemId = c.StoreItemId!,
                })
                .ToList();

        // ---- new unbound fields: a genuine OLD-vs-NEW diff (non-label, new to this schema, still unbound) ----
        //
        // Empty when the old part could not be read: with no old schema there is no diff to compute, and
        // listing every unbound column instead would report the fields the layout simply never used as "new"
        // (see RefreshResult.NewUnboundFields).
        var boundPaths = BuildBoundPaths(inventory);
        var newUnboundFields = target.Unreadable is not null
            ? []
            : newTree.AllColumns(includeSystem: false)
                .Where(c => !c.IsLabel && !oldColumnPaths.Contains(c.Path) && !boundPaths.Contains(c.Path))
                .Select(c => c.Path)
                .ToList();

        // ---- mutate: replace the BC part's raw content, keeping the SAME part / the SAME ds:itemID ----
        using (var partStream = bcPart.GetStream(FileMode.Create, FileAccess.Write))
        {
            partStream.Write(newBytes, 0, newBytes.Length);
        }

        UpdateSchemaReference(bcPart, newIdentity.Namespace);

        // ---- re-point EVERY BC-namespaced binding at the new namespace, and every BC tag at the new report ----
        //
        // Unconditional, not gated on namespaceChanged: a binding can name a namespace that is neither the old
        // nor the new one (20 of PaymentPracticeByPeriod.docx's 25 do), and those are broken whether or not
        // THIS refresh changed the report's own namespace. Gating on namespaceChanged left them untouched
        // whenever the refresh was a same-namespace one, which is the common case (GitHub issue #2).
        RepointToNewIdentity(main, newIdentity);
        if (oldStoreItemId is not null && repointedStoreItemIds.Count > 0)
        {
            RepointDanglingStoreItemIds(main, oldStoreItemId, partItemIds);
        }

        var quick = LayoutValidator.Quick(doc);

        return new RefreshResult
        {
            OldReportName = oldIdentity.ReportName,
            OldReportId = oldIdentity.ReportId,
            OldNamespace = oldIdentity.Namespace,
            NewReportName = newIdentity.ReportName,
            NewReportId = newIdentity.ReportId,
            NewNamespace = newIdentity.Namespace,
            StoreItemId = oldStoreItemId,
            NamespaceChanged = namespaceChanged,
            RemappedCount = remappedCount,
            OrphanedBindings = orphaned,
            RepointedForeignBindings = repointedForeign,
            RepointedStoreItemIds = repointedStoreItemIds,
            NewUnboundFields = newUnboundFields,
            RepairedUnreadablePart = target.Unreadable?.Describe(),
            QuickValidation = quick,
        };
    }

    // ---- the part to refresh: the layout's readable BC part, or the one unreadable candidate to repair ----

    /// <summary>
    /// The BC part a refresh replaces, plus what is known of it BEFORE its content is overwritten: its old
    /// identity and its old leaf column paths (empty, with <see cref="Unreadable"/> set, when it cannot be read).
    /// </summary>
    private sealed record RefreshTarget(
        CustomXmlPart Part, ReportIdentity OldIdentity, IReadOnlySet<string> OldColumnPaths, UnreadableBcPart? Unreadable);

    /// <summary>
    /// Picks the part this refresh replaces. Ordinarily the layout's readable BC dataset part, whose OLD
    /// schema is read in full first - <see cref="RefreshResult.NewUnboundFields"/> needs the old leaf column
    /// paths to report a genuine OLD-vs-NEW diff rather than every unbound leaf still sitting in the new schema.
    /// </summary>
    /// <remarks>
    /// REPAIR (GitHub issue #96). When no part parses as a BC dataset part but exactly one is recognisably
    /// one (<see cref="SchemaProvider.Scan"/>), that part is the target: replacing its content is exactly the
    /// repair it needs, and this is the only tool that can make it - every other tool needs the part readable
    /// first. Its old identity is what its itemProps schemaRef or raw bytes name (all that can be known of an
    /// unparsable part), and it keeps its <c>ds:itemID</c> like any refreshed part, so the layout's bindings stay
    /// linked. Refusing here, as this used to ("Layout has no BC dataset custom XML part to refresh"), left the
    /// layout with no repair path inside the tool set at all.
    /// <para>
    /// SEVERAL BC parts are refused (GitHub issue #101): this used to replace "the first" in relationship
    /// order, so with a foreign part listed first it rewrote THAT part, reported namespaceChanged, listed
    /// dozens of the layout's own columns as new, and left errors no tool could clear.
    /// <see cref="SchemaProvider.FindBcPart"/> is the refusal, unreadable candidates counting towards it.
    /// </para>
    /// </remarks>
    private static RefreshTarget ResolveTarget(MainDocumentPart main)
    {
        var scan = SchemaProvider.Scan(main);
        if (scan.Count == 1 && scan.Unreadable is [var broken])
        {
            return new RefreshTarget(
                broken.Part,
                SchemaProvider.ParseIdentity(broken.Namespace, broken.ItemId),
                new HashSet<string>(StringComparer.Ordinal),
                broken);
        }

        var (bcPart, bcRoot) = SchemaProvider.FindBcPart(main)
            ?? throw new InvalidDataException(
                $"Layout has no BC dataset custom XML part (namespace starting '{OoxmlNames.BcNamespacePrefix}') to refresh.");

        var oldColumnPaths = SchemaProvider.FromMainPart(main).AllColumns(includeSystem: false)
            .Select(c => c.Path)
            .ToHashSet(StringComparer.Ordinal);

        return new RefreshTarget(
            bcPart,
            SchemaProvider.ParseIdentity(bcRoot.Name.NamespaceName, SchemaProvider.ReadItemId(bcPart)),
            oldColumnPaths,
            Unreadable: null);
    }

    // ---- new schema loading (.docx layout vs standalone schema .xml) ----

    /// <summary>
    /// Loads the raw new-schema bytes plus its fully parsed <see cref="DatasetTree"/> from
    /// <paramref name="newSchemaSource"/>. Mirrors <c>LayoutBuilder.LoadDatasetSource</c>'s exact split
    /// (<c>.docx</c> layout vs standalone schema <c>.xml</c>) and raw-bytes/encoding handling, but returns
    /// the full tree rather than only the identity - <see cref="Refresh"/> needs the whole schema shape to
    /// classify existing bindings and to enumerate newly unbound fields, not just the report name/id.
    /// </summary>
    private static (byte[] Bytes, DatasetTree Tree) LoadNewSchema(string newSchemaSource)
    {
        if (!File.Exists(newSchemaSource))
        {
            throw new FileNotFoundException("newSchemaSource does not point to an existing file.", newSchemaSource);
        }

        if (newSchemaSource.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
        {
            using var sourceDoc = LayoutPackage.Open(newSchemaSource, isEditable: false);
            var sourceMain = sourceDoc.MainDocumentPart
                ?? throw new InvalidDataException($"'{newSchemaSource}' has no main document part.");

            // FindBcPart refuses a source carrying several BC parts too: which schema to copy would be a guess.
            var found = SchemaProvider.FindBcPart(sourceMain, $"newSchemaSource '{newSchemaSource}'")
                ?? throw new InvalidDataException(
                    $"'{newSchemaSource}' has no BC dataset custom XML part (namespace starting "
                    + $"'{OoxmlNames.BcNamespacePrefix}').");

            var tree = SchemaProvider.FromLayout(sourceDoc);

            using var partStream = found.Part.GetStream(FileMode.Open, FileAccess.Read);
            var bytes = ResourceLimits.ReadAllBytesCapped(
                partStream, $"Custom XML part '{PartWalker.PartFileName(found.Part)}' in '{newSchemaSource}'");
            return (bytes, tree);
        }

        var schemaTree = SchemaProvider.FromSchemaXml(newSchemaSource);
        return (File.ReadAllBytes(newSchemaSource), schemaTree);
    }

    // ---- bound-path set (mirrors BcWordLayout.McpHost.Tools.ToolGuards.BuildBoundPaths) ----

    /// <summary>
    /// Converts every control's binding XPath into a <see cref="DatasetColumn.Path"/>-comparable string
    /// (drop the root element-name segment, rejoin the rest with a leading slash), so it can be checked
    /// against the new schema's own leaf column paths.
    /// </summary>
    private static HashSet<string> BuildBoundPaths(LayoutInventory inventory)
    {
        var set = new HashSet<string>(StringComparer.Ordinal);
        foreach (var control in inventory.Controls)
        {
            if (control.XPath is null)
            {
                continue;
            }

            var segments = BindingXPath.Segments(control.XPath);
            if (segments.Count <= 1)
            {
                continue;
            }

            set.Add("/" + string.Join("/", segments.Skip(1)));
        }

        return set;
    }

    // ---- BC part content swap: keep the same part / ds:itemID, refresh the schemaRef URI ----

    /// <summary>
    /// Updates the BC part's <c>ds:schemaRefs/ds:schemaRef/@ds:uri</c> to <paramref name="newNamespace"/> so
    /// it stays consistent with the part's just-replaced content (harmless, and a no-op value-wise when the
    /// namespace did not change). A no-op when the part has no schema reference to update (e.g. a minimal
    /// synthetic/test fixture) rather than an error - this is a best-effort consistency touch-up, not part of
    /// what makes the refresh itself correct (BC/<see cref="SchemaProvider"/> both key off the part's actual
    /// root namespace, not this reference).
    /// </summary>
    private static void UpdateSchemaReference(CustomXmlPart bcPart, string newNamespace)
    {
        var dataStoreItem = bcPart.CustomXmlPropertiesPart?.DataStoreItem;
        var schemaRef = dataStoreItem?.SchemaReferences?.Elements<Ds.SchemaReference>().FirstOrDefault();
        if (schemaRef is null)
        {
            return;
        }

        schemaRef.Uri = newNamespace;
        dataStoreItem!.Save();
    }

    // ---- namespace-change remap: w:prefixMappings URI (binding) + w:tag (every BC-authored control) ----

    /// <summary>
    /// Walks <paramref name="main"/>'s document plus every header and footer part (the same three-part scope
    /// <see cref="LayoutReader.Read(WordprocessingDocument)"/> covers), re-pointing every binding's
    /// <c>w:prefixMappings</c> URI at <paramref name="newIdentity"/>'s namespace and every BC-authored
    /// control's <c>w:tag</c> at its <c>#Nav: &lt;ReportName&gt;/&lt;ReportId&gt;</c> form. The XPath itself is
    /// never touched here - its element-name steps are what "remap where element names match" means, and they
    /// do not depend on the report's own name/id.
    /// </summary>
    /// <remarks>
    /// Re-points by WHATEVER BC namespace a binding currently names, rather than by matching one specific old
    /// namespace. That is the whole of GitHub issue #2: the previous implementation swapped a known
    /// old-for-new substring, so a binding naming a THIRD namespace (a superseded namespace of the same
    /// report, or another report's entirely - 20 of PaymentPracticeByPeriod.docx's 25 bindings) survived every
    /// refresh untouched, and the only repair on offer was to delete and rebuild each control by hand. The
    /// 2026-08-02 sandbox rounds (issue #1) settled that such a binding is not merely unusual: BC validates
    /// every binding's prefixMappings against the target report's CURRENT namespace and refuses the upload
    /// with one <c>InvalidPrefixMapping</c> per offender, accepting the same layout the moment they are all
    /// re-pointed. So re-pointing is not a guess about caller intent (ADR-0003) - it is the only state BC
    /// accepts, and the alternative was leaving the layout provably un-uploadable.
    /// <para>
    /// A binding whose <c>prefixMappings</c> names no BC namespace at all is left exactly as found (see
    /// <see cref="RepointNamespace"/>): the point is to make BC-bound controls name THIS report, not to claim
    /// every content control in the document for it.
    /// </para>
    /// </remarks>
    private static void RepointToNewIdentity(MainDocumentPart main, ReportIdentity newIdentity)
    {
        var newTag = $"#Nav: {newIdentity.ReportName}/{newIdentity.ReportId}";

        foreach (var (root, _) in PartWalker.ContentParts(main))
        {
            RepointPart(root, newIdentity.Namespace, newTag);
        }
    }

    /// <summary>
    /// Points every BC-namespaced binding whose <c>w:storeItemID</c> names none of <paramref name="partItemIds"/>
    /// at <paramref name="bcItemId"/>, in the same three-part scope as <see cref="RepointToNewIdentity"/>.
    /// Runs after it, so every BC binding already names the new namespace.
    /// </summary>
    private static void RepointDanglingStoreItemIds(MainDocumentPart main, string bcItemId, IReadOnlyList<string> partItemIds)
    {
        bool Dangling(string? storeItemId) =>
            storeItemId is not null && !partItemIds.Any(id => SchemaProvider.SameItemId(id, storeItemId));

        foreach (var (root, _) in PartWalker.ContentParts(main))
        {
            foreach (var pr in root.Descendants<SdtProperties>())
            {
                switch (SdtInspector.FindBinding(pr))
                {
                    case DataBinding w when SdtInspector.ExtractBcNamespace(w.PrefixMappings?.Value) is not null
                                            && Dangling(w.StoreItemId?.Value):
                        w.StoreItemId = bcItemId;
                        break;

                    case Office2013Word.DataBinding w15 when SdtInspector.ExtractBcNamespace(w15.PrefixMappings?.Value) is not null
                                                            && Dangling(w15.StoreItemId?.Value):
                        w15.StoreItemId = bcItemId;
                        break;
                }
            }
        }
    }

    private static void RepointPart(OpenXmlElement root, string newNamespace, string newTag)
    {
        foreach (var pr in root.Descendants<SdtProperties>())
        {
            RepointBindingPrefix(pr, newNamespace);
            RemapTag(pr, newTag);
        }
    }

    /// <summary>
    /// Rewrites the BC URI inside <c>w:prefixMappings="xmlns:ns0='&lt;uri&gt;'"</c> (or the
    /// <c>w15:dataBinding</c> equivalent) to <paramref name="newNamespace"/>.
    /// </summary>
    /// <remarks>
    /// Assigns only when there is a rewrite to make. Writing the unchanged value back is NOT a no-op: for a
    /// binding that carries no <c>w:prefixMappings</c> at all — legal, and present in real layouts — assigning
    /// null through <c>StringValue</c> materialises an EMPTY <c>w:prefixMappings=""</c> attribute where there
    /// was none. That was latent in the old old-for-new implementation too (reachable on any
    /// namespace-changing refresh); making the re-point unconditional would have made it routine.
    /// </remarks>
    private static void RepointBindingPrefix(SdtProperties pr, string newNamespace)
    {
        var binding = SdtInspector.FindBinding(pr);
        var xpath = binding is null ? null : SdtInspector.Attr(binding, "xpath", OoxmlNames.W);

        switch (binding)
        {
            case DataBinding w when RepointNamespace(w.PrefixMappings?.Value, xpath, newNamespace) is { } updated:
                w.PrefixMappings = updated;
                break;

            case Office2013Word.DataBinding w15
                when RepointNamespace(w15.PrefixMappings?.Value, xpath, newNamespace) is { } updated:
                w15.PrefixMappings = updated;
                break;
        }
    }

    /// <summary>
    /// Replaces the BC namespace <paramref name="prefixMappings"/> currently names with
    /// <paramref name="newNamespace"/>, leaving everything else in an <c>xmlns:</c> declaration list
    /// byte-for-byte: the prefix itself (<c>ns0</c>), the quoting, and the incidental formatting real corpus
    /// bindings are inconsistent about (a trailing space before the closing quote). Finding the current URI
    /// via <see cref="SdtInspector.ExtractBcNamespace"/> - the same parser the read/validate side uses to
    /// REPORT a binding's namespace - is what makes that possible without knowing the old value, and keeps
    /// write and read agreeing by construction.
    /// <para>
    /// A BARE URI with no <c>xmlns:</c> declaration at all (StandardSalesInvoiceVatSpec.docx's
    /// Mini_Sales_Invoice bindings) is the one exception: it is rewritten as
    /// <c>xmlns:&lt;prefix&gt;='&lt;newNamespace&gt;'</c> - the form <see cref="SdtFactory"/> writes - for the
    /// single prefix <paramref name="xpath"/> uses. Swapping only the URI left it bare, and Word reports a
    /// bare-mapped control as unmapped (ContentControl.XMLMapping.IsMapped is False on both re-pointed
    /// bindings, GitHub issue #79), so the repair stopped halfway. Writing the declaration is not a guess:
    /// the XPath names the prefix and the re-point fixes the URI it must map to. An XPath with no prefix, or
    /// with several, gives no single prefix to declare, so such a value keeps the plain URI swap.
    /// </para>
    /// <para>
    /// Returns <c>null</c> to mean "nothing to rewrite" — the value names no BC namespace (so it is none of
    /// this method's business) or already names the new one — which is what lets the caller leave such a
    /// binding's attribute completely alone; see <see cref="RepointBindingPrefix"/>'s remarks for why that
    /// distinction matters. A bare mapping that already names the new namespace is therefore left bare: this
    /// method repairs what it re-points, and does not rewrite a binding BC already accepts.
    /// </para>
    /// </summary>
    private static string? RepointNamespace(string? prefixMappings, string? xpath, string newNamespace)
    {
        var current = SdtInspector.ExtractBcNamespace(prefixMappings);
        if (current is null || string.Equals(current, newNamespace, StringComparison.Ordinal))
        {
            return null;
        }

        if (BindingXPath.PrefixDeclarations(prefixMappings).Count == 0 && BindingXPath.Prefixes(xpath) is [var prefix])
        {
            return $"xmlns:{prefix}='{newNamespace}'";
        }

        return prefixMappings!.Replace(current, newNamespace, StringComparison.Ordinal);
    }

    /// <summary>
    /// Rewrites <paramref name="pr"/>'s <c>w:tag</c> to <paramref name="newTag"/>, but only when a tag is
    /// present AND already follows the BC <c>#Nav: ...</c> convention - a defensive filter so an unrelated
    /// content control that happens to carry some other <c>w:tag</c> value is never touched.
    /// </summary>
    private static void RemapTag(SdtProperties pr, string newTag)
    {
        var tag = pr.GetFirstChild<Tag>();
        if (tag?.Val?.Value?.StartsWith("#Nav:", StringComparison.Ordinal) == true)
        {
            tag.Val = newTag;
        }
    }

}
