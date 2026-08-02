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
    /// <paramref name="doc"/> has no main document part or no BC dataset custom XML part to refresh;
    /// <paramref name="newSchemaSource"/> is a <c>.docx</c> with no main document part or no BC dataset
    /// custom XML part; or is a schema <c>.xml</c> whose root is not <c>NavWordReportXmlPart</c>.
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
        var (bcPart, bcRoot) = SchemaProvider.FindBcPart(main)
            ?? throw new InvalidDataException(
                $"Layout has no BC dataset custom XML part (namespace starting '{OoxmlNames.BcNamespacePrefix}') to refresh.");

        var oldStoreItemId = bcPart.CustomXmlPropertiesPart?.DataStoreItem?.ItemId?.Value;
        var oldIdentity = SchemaProvider.ParseIdentity(bcRoot.Name.NamespaceName, oldStoreItemId);

        // Capture the OLD schema's own leaf column paths BEFORE the BC part's content is overwritten below -
        // NewUnboundFields (see its own remarks) needs this to report a genuine OLD-vs-NEW diff rather than
        // every unbound leaf that happens to still be sitting in the new schema.
        var oldColumnPaths = SchemaProvider.FromMainPart(main).AllColumns(includeSystem: false)
            .Select(c => c.Path)
            .ToHashSet(StringComparer.Ordinal);

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

        // ---- new unbound fields: a genuine OLD-vs-NEW diff (non-label, new to this schema, still unbound) ----
        var boundPaths = BuildBoundPaths(inventory);
        var newUnboundFields = newTree.AllColumns(includeSystem: false)
            .Where(c => !c.IsLabel && !oldColumnPaths.Contains(c.Path) && !boundPaths.Contains(c.Path))
            .Select(c => c.Path)
            .ToList();

        // ---- mutate: replace the BC part's raw content, keeping the SAME part / the SAME ds:itemID ----
        using (var partStream = bcPart.GetStream(FileMode.Create, FileAccess.Write))
        {
            partStream.Write(newBytes, 0, newBytes.Length);
        }

        UpdateSchemaReference(bcPart, newIdentity.Namespace);

        // ---- namespace change: remap every binding's prefixMappings URI + every control's tag ----
        if (namespaceChanged)
        {
            RemapNamespace(main, oldIdentity.Namespace, newIdentity);
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
            NewUnboundFields = newUnboundFields,
            QuickValidation = quick,
        };
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
            using var sourceDoc = WordprocessingDocument.Open(newSchemaSource, false);
            var sourceMain = sourceDoc.MainDocumentPart
                ?? throw new InvalidDataException($"'{newSchemaSource}' has no main document part.");

            var found = SchemaProvider.FindBcPart(sourceMain)
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
    /// <see cref="LayoutReader.Read(WordprocessingDocument)"/> covers) rewriting every binding's
    /// <c>w:prefixMappings</c> URI from <paramref name="oldNamespace"/> to <paramref name="newIdentity"/>'s
    /// namespace, and every BC-authored control's <c>w:tag</c> to <paramref name="newIdentity"/>'s
    /// <c>#Nav: &lt;ReportName&gt;/&lt;ReportId&gt;</c> form. The XPath itself is never touched here - its
    /// element-name steps are what "remap where element names match" means, and they do not depend on the
    /// report's own name/id.
    /// </summary>
    private static void RemapNamespace(MainDocumentPart main, string oldNamespace, ReportIdentity newIdentity)
    {
        var newTag = $"#Nav: {newIdentity.ReportName}/{newIdentity.ReportId}";

        foreach (var (root, _) in PartWalker.ContentParts(main))
        {
            RemapPart(root, oldNamespace, newIdentity.Namespace, newTag);
        }
    }

    private static void RemapPart(OpenXmlElement root, string oldNamespace, string newNamespace, string newTag)
    {
        foreach (var pr in root.Descendants<SdtProperties>())
        {
            RemapBindingPrefix(pr, oldNamespace, newNamespace);
            RemapTag(pr, newTag);
        }
    }

    /// <summary>
    /// Rewrites the URI inside <c>w:prefixMappings="xmlns:ns0='&lt;uri&gt;'"</c> (or the <c>w15:dataBinding</c>
    /// equivalent) via a targeted substring replace of the OLD namespace for the NEW one - the prefix itself
    /// (<c>ns0</c>) and any incidental formatting (real corpus bindings are inconsistent about a trailing
    /// space before the closing quote) are left exactly as found, matching "prefix stays ns0; only the
    /// mapped URI changes".
    /// </summary>
    private static void RemapBindingPrefix(SdtProperties pr, string oldNamespace, string newNamespace)
    {
        switch (SdtInspector.FindBinding(pr))
        {
            case DataBinding w:
                w.PrefixMappings = ReplaceNamespace(w.PrefixMappings?.Value, oldNamespace, newNamespace);
                break;

            case Office2013Word.DataBinding w15:
                w15.PrefixMappings = ReplaceNamespace(w15.PrefixMappings?.Value, oldNamespace, newNamespace);
                break;
        }
    }

    private static string? ReplaceNamespace(string? prefixMappings, string oldNamespace, string newNamespace) =>
        prefixMappings is not null && prefixMappings.Contains(oldNamespace, StringComparison.Ordinal)
            ? prefixMappings.Replace(oldNamespace, newNamespace, StringComparison.Ordinal)
            : prefixMappings;

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
