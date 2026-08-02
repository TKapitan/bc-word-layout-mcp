using System.Xml.Linq;
using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;
using DocumentFormat.OpenXml.Packaging;

namespace BcWordLayout.Tests;

/// <summary>
/// Covers <see cref="LayoutRefresher"/> directly (no MCP tool layer): each test opens a temp COPY of a
/// corpus layout editable, refreshes it against either its OWN schema (idempotent case) or a MODIFIED copy
/// of that schema built by editing the corpus's own raw BC-part XML (add/remove/rename a column, or change
/// the report id in the namespace) — mirroring the round-trip style already used by
/// <c>LayoutBuilderTests</c>/<c>LayoutEditorTests</c>.
/// </summary>
public class LayoutRefresherTests
{
    // Real, verified pre-existing bound controls in tests/corpus/SalesInvoiceForSubscriptionBilling.docx
    // (Standard_Sales_Invoice/1306) - both direct /Header/<name> leaf columns (see SdtFactory's own remarks
    // doc-comment and LayoutEditorTests' CellLevelLabelId, which already rely on these same two controls).
    private const string FieldColumn = "SalesPersonName"; // a field control, /Header/SalesPersonName
    private const string LabelColumn = "YourReference_Lbl"; // a label control, /Header/YourReference_Lbl

    // ---- fixture plumbing: temp copies, and modified schema .xml files built from the corpus's own BC part ----

    private static string CopyOfCorpus(string corpusFile)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-refresher-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    /// <summary>
    /// Finds every custom XML part whose root namespace starts with the BC prefix, via PUBLIC API only
    /// (mirrors <c>LayoutBuilderTests.FindBcCustomXmlParts</c> - <c>SchemaProvider.FindBcParts</c> is
    /// internal and not visible from this test assembly).
    /// </summary>
    private static List<CustomXmlPart> FindBcCustomXmlParts(MainDocumentPart main)
    {
        var result = new List<CustomXmlPart>();
        foreach (var part in main.CustomXmlParts)
        {
            using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
            XElement? root;
            try
            {
                root = XDocument.Load(stream).Root;
            }
            catch
            {
                continue;
            }

            if (root is not null && root.Name.NamespaceName.StartsWith(OoxmlNames.BcNamespacePrefix, StringComparison.Ordinal))
            {
                result.Add(part);
            }
        }

        return result;
    }

    private static XDocument LoadCorpusSchemaXDocument(string corpusDocxPath)
    {
        using var doc = WordprocessingDocument.Open(corpusDocxPath, false);
        var bcPart = FindBcCustomXmlParts(doc.MainDocumentPart!).Single();
        using var stream = bcPart.GetStream(FileMode.Open, FileAccess.Read);
        return XDocument.Load(stream);
    }

    private static string SaveAsTempSchemaXml(XDocument xdoc)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-refresher-schema-{Guid.NewGuid():N}.xml");
        xdoc.Save(path);
        return path;
    }

    /// <summary>Builds a modified schema .xml with <paramref name="columnLocalName"/> removed from under <paramref name="parentLocalName"/>.</summary>
    private static string BuildSchemaWithColumnRemoved(string corpusDocxPath, string parentLocalName, string columnLocalName)
    {
        var xdoc = LoadCorpusSchemaXDocument(corpusDocxPath);
        var ns = xdoc.Root!.Name.Namespace;
        var parent = xdoc.Root.Element(ns + parentLocalName)!;
        parent.Element(ns + columnLocalName)!.Remove();
        return SaveAsTempSchemaXml(xdoc);
    }

    /// <summary>Builds a modified schema .xml with <paramref name="oldName"/> renamed to <paramref name="newName"/> under <paramref name="parentLocalName"/>.</summary>
    private static string BuildSchemaWithColumnRenamed(string corpusDocxPath, string parentLocalName, string oldName, string newName)
    {
        var xdoc = LoadCorpusSchemaXDocument(corpusDocxPath);
        var ns = xdoc.Root!.Name.Namespace;
        var parent = xdoc.Root.Element(ns + parentLocalName)!;
        var column = parent.Element(ns + oldName)!;
        column.Name = ns + newName;
        column.Value = newName;
        return SaveAsTempSchemaXml(xdoc);
    }

    /// <summary>Builds a modified schema .xml with a brand-new leaf column added under <paramref name="parentLocalName"/>.</summary>
    private static string BuildSchemaWithColumnAdded(string corpusDocxPath, string parentLocalName, string newColumnName)
    {
        var xdoc = LoadCorpusSchemaXDocument(corpusDocxPath);
        var ns = xdoc.Root!.Name.Namespace;
        var parent = xdoc.Root.Element(ns + parentLocalName)!;
        parent.Add(new XElement(ns + newColumnName, newColumnName));
        return SaveAsTempSchemaXml(xdoc);
    }

    /// <summary>
    /// Builds a modified schema .xml whose EVERY element (root + all descendants) moves to
    /// <paramref name="newNamespace"/> - simulating a real AL report id bump, where the whole exported
    /// dataset lives in one default <c>xmlns</c> declaration at the root, so every element is affected.
    /// </summary>
    private static string BuildSchemaWithNewNamespace(string corpusDocxPath, string newNamespace)
    {
        var xdoc = LoadCorpusSchemaXDocument(corpusDocxPath);
        var ns = XNamespace.Get(newNamespace);

        // Materialize BEFORE mutating: renaming an element while a live Descendants() query is still
        // enumerating the same tree is unsafe (it raises a Name-changed event mid-walk), so snapshot the
        // full element list first, then rename every one of them.
        var elements = xdoc.Root!.DescendantsAndSelf().ToList();
        foreach (var el in elements)
        {
            // XDocument.Load materializes the root's own "xmlns='...'" declaration as an explicit
            // XAttribute (IsNamespaceDeclaration). Renaming .Name alone does not touch it, so the STALE
            // declaration (still pointing at the OLD namespace) would conflict with the new one when
            // saving ("prefix '' cannot be redefined ... within the same start element tag"). Drop it and
            // let the writer regenerate whatever declaration is actually needed from the updated names.
            foreach (var declaration in el.Attributes().Where(a => a.IsNamespaceDeclaration).ToList())
            {
                declaration.Remove();
            }

            el.Name = ns + el.Name.LocalName;
        }

        return SaveAsTempSchemaXml(xdoc);
    }

    private static int CountBoundControls(WordprocessingDocument doc) =>
        LayoutReader.Read(doc).Controls.Count(c => c.XPath is not null);

    private static void SaveAllParts(WordprocessingDocument doc)
    {
        doc.MainDocumentPart!.Document!.Save();
        foreach (var header in doc.MainDocumentPart.HeaderParts)
        {
            header.Header?.Save();
        }

        foreach (var footer in doc.MainDocumentPart.FooterParts)
        {
            footer.Footer?.Save();
        }
    }

    // ---- idempotent: refresh with the layout's own, unmodified schema ----

    [Fact]
    public void Refresh_with_its_own_unmodified_schema_is_idempotent()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            RefreshResult result;
            int boundBefore;
            string? storeItemIdBefore;
            int dataItemsBefore;
            int columnsBefore;

            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var before = SchemaProvider.FromLayout(doc);
                storeItemIdBefore = before.Report.StoreItemId;
                dataItemsBefore = before.AllDataItems(includeSystem: true).Count();
                columnsBefore = before.AllColumns(includeSystem: true).Count();
                boundBefore = CountBoundControls(doc);

                result = LayoutRefresher.Refresh(doc, Corpus.Path(Corpus.SalesInvoice));

                SaveAllParts(doc);
            }

            Assert.False(result.NamespaceChanged);
            Assert.Equal(result.OldNamespace, result.NewNamespace);
            Assert.Empty(result.OrphanedBindings);
            Assert.Equal(boundBefore, result.RemappedCount);
            Assert.Equal(storeItemIdBefore, result.StoreItemId, StringComparer.OrdinalIgnoreCase);

            // NewUnboundFields is an OLD-vs-NEW diff: refreshing against the layout's OWN unchanged schema
            // introduces no new columns at all, so this MUST be empty - not merely "a subset of real
            // columns" (that weaker check is exactly what let the old all-unbound-leaves bug through: a
            // refresh against an unchanged schema used to report every pre-existing unbound field, e.g. 171
            // of them on this corpus layout, burying any genuinely new one).
            Assert.Empty(result.NewUnboundFields);

            Assert.Equal("quick", result.QuickValidation.Level);
            Assert.True(result.QuickValidation.Passed,
                "expected a clean refresh; errors: " + string.Join(" | ",
                    result.QuickValidation.Findings.Where(f => f.Severity == FindingSeverity.Error).Select(f => f.Message)));
            Assert.Equal(0, result.QuickValidation.ErrorCount);

            using var reopened = WordprocessingDocument.Open(path, false);
            var after = SchemaProvider.FromLayout(reopened);
            Assert.Equal(storeItemIdBefore, after.Report.StoreItemId, StringComparer.OrdinalIgnoreCase);
            Assert.Equal(dataItemsBefore, after.AllDataItems(includeSystem: true).Count());
            Assert.Equal(columnsBefore, after.AllColumns(includeSystem: true).Count());
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- removed column -> orphaned; a still-valid sibling is not ----

    [Fact]
    public void Refresh_reports_orphaned_binding_for_a_removed_column_but_not_for_a_still_valid_sibling()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        string? schemaPath = null;
        try
        {
            schemaPath = BuildSchemaWithColumnRemoved(Corpus.Path(Corpus.SalesInvoice), "Header", FieldColumn);

            RefreshResult result;
            int boundBefore;
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                boundBefore = CountBoundControls(doc);
                result = LayoutRefresher.Refresh(doc, schemaPath);
                SaveAllParts(doc);
            }

            Assert.False(result.NamespaceChanged);
            Assert.True(result.OrphanedBindings.Count >= 1, "expected at least one orphaned binding");
            Assert.All(result.OrphanedBindings, o => Assert.EndsWith($"{FieldColumn}[1]", o.XPath, StringComparison.Ordinal));
            Assert.DoesNotContain(result.OrphanedBindings, o => o.XPath.Contains(LabelColumn, StringComparison.Ordinal));
            Assert.Equal(boundBefore, result.RemappedCount + result.OrphanedBindings.Count);

            // The post-refresh quick validation independently surfaces the exact same fact.
            Assert.False(result.QuickValidation.Passed);
            Assert.Contains(result.QuickValidation.Findings, f =>
                f.Check == "xpath-resolves" && (f.Location ?? "").Contains(FieldColumn, StringComparison.Ordinal));
        }
        finally
        {
            File.Delete(path);
            if (schemaPath is not null)
            {
                File.Delete(schemaPath);
            }
        }
    }

    // ---- orphan report carries the control's w:id so the caller can remove_control it directly ----

    [Fact]
    public void Refresh_orphaned_binding_exposes_the_control_sdtId_for_direct_removal()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        string? schemaPath = null;
        try
        {
            schemaPath = BuildSchemaWithColumnRemoved(Corpus.Path(Corpus.SalesInvoice), "Header", FieldColumn);

            RefreshResult result;
            HashSet<int> inventoryIdsBefore;
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                // Every w:id present in the layout BEFORE the refresh (refresh never changes ids).
                inventoryIdsBefore = LayoutReader.Read(doc).Controls
                    .Where(c => c.SdtId is not null)
                    .Select(c => c.SdtId!.Value)
                    .ToHashSet();

                result = LayoutRefresher.Refresh(doc, schemaPath);
                SaveAllParts(doc);
            }

            Assert.NotEmpty(result.OrphanedBindings);

            // The whole point: an orphan report must hand back a usable w:id (remove_control requires it),
            // and that id must name a control that actually exists in the layout - no second get_layout_info
            // round trip, no ambiguous alias/XPath match needed to remediate.
            Assert.All(result.OrphanedBindings, o =>
            {
                Assert.NotNull(o.SdtId);
                Assert.Contains(o.SdtId!.Value, inventoryIdsBefore);
            });
        }
        finally
        {
            File.Delete(path);
            if (schemaPath is not null)
            {
                File.Delete(schemaPath);
            }
        }
    }

    // ---- renamed column -> orphaned; a still-valid sibling is not ----

    [Fact]
    public void Refresh_reports_orphaned_binding_for_a_renamed_column_but_not_for_a_still_valid_sibling()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        string? schemaPath = null;
        try
        {
            schemaPath = BuildSchemaWithColumnRenamed(
                Corpus.Path(Corpus.SalesInvoice), "Header", LabelColumn, LabelColumn + "Renamed");

            RefreshResult result;
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                result = LayoutRefresher.Refresh(doc, schemaPath);
                SaveAllParts(doc);
            }

            Assert.True(result.OrphanedBindings.Count >= 1, "expected at least one orphaned binding");
            Assert.All(result.OrphanedBindings, o => Assert.EndsWith($"{LabelColumn}[1]", o.XPath, StringComparison.Ordinal));
            Assert.DoesNotContain(result.OrphanedBindings, o => o.XPath.Contains(FieldColumn, StringComparison.Ordinal));
        }
        finally
        {
            File.Delete(path);
            if (schemaPath is not null)
            {
                File.Delete(schemaPath);
            }
        }
    }

    // ---- added column -> new unbound field; storeItemID preserved; part content really replaced ----

    [Fact]
    public void Refresh_reports_a_new_unbound_field_for_an_added_column_preserves_storeItemId_and_replaces_part_content()
    {
        const string newColumn = "BrandNewFieldXYZ123";
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        string? schemaPath = null;
        try
        {
            schemaPath = BuildSchemaWithColumnAdded(Corpus.Path(Corpus.SalesInvoice), "Header", newColumn);

            RefreshResult result;
            string? storeItemIdBefore;
            int boundBefore;
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                storeItemIdBefore = SchemaProvider.FromLayout(doc).Report.StoreItemId;
                boundBefore = CountBoundControls(doc);

                result = LayoutRefresher.Refresh(doc, schemaPath);

                SaveAllParts(doc);
            }

            // Exactly the newly-added field - not a tolerant subset check. Since this is an OLD-vs-NEW diff,
            // every OTHER non-label leaf column (whether bound or not) already existed in the OLD schema too
            // and must NOT resurface here, however many of them remain genuinely unbound.
            Assert.Equal(new[] { "/Header/" + newColumn }, result.NewUnboundFields);
            Assert.Empty(result.OrphanedBindings);
            Assert.Equal(boundBefore, result.RemappedCount);
            Assert.Equal(storeItemIdBefore, result.StoreItemId, StringComparer.OrdinalIgnoreCase);
            Assert.NotNull(storeItemIdBefore);

            using var reopened = WordprocessingDocument.Open(path, false);
            var after = SchemaProvider.FromLayout(reopened);

            // storeItemID survives a real disk round-trip unchanged.
            Assert.Equal(storeItemIdBefore, after.Report.StoreItemId, StringComparer.OrdinalIgnoreCase);

            // The part's CONTENT was really replaced: the new column is now present in the schema tree.
            Assert.Contains(after.AllColumns(includeSystem: false), c => c.Name == newColumn && c.Path == "/Header/" + newColumn);
        }
        finally
        {
            File.Delete(path);
            if (schemaPath is not null)
            {
                File.Delete(schemaPath);
            }
        }
    }

    // ---- namespace change: prefixMappings/tag remap; bindings still resolve by element name ----

    [Fact]
    public void Refresh_with_a_different_report_id_remaps_prefixMappings_and_tags_and_bindings_still_resolve()
    {
        const string newNamespace = "urn:microsoft-dynamics-nav/reports/Standard_Sales_Invoice/9999/";
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        string? schemaPath = null;
        try
        {
            schemaPath = BuildSchemaWithNewNamespace(Corpus.Path(Corpus.SalesInvoice), newNamespace);

            RefreshResult result;
            int boundBefore;
            string oldNamespace;
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                oldNamespace = SchemaProvider.FromLayout(doc).Report.Namespace;
                boundBefore = CountBoundControls(doc);

                result = LayoutRefresher.Refresh(doc, schemaPath);

                SaveAllParts(doc);
            }

            Assert.True(result.NamespaceChanged);
            Assert.Equal("Standard_Sales_Invoice", result.NewReportName);
            Assert.Equal("9999", result.NewReportId);
            Assert.Equal(newNamespace, result.NewNamespace);
            Assert.Equal(oldNamespace, result.OldNamespace);
            Assert.NotEqual(result.OldNamespace, result.NewNamespace, StringComparer.Ordinal);

            // Same element names throughout -> nothing is orphaned, everything remaps, and no column paths
            // changed either (only the namespace did), so the OLD-vs-NEW diff is empty too.
            Assert.Empty(result.OrphanedBindings);
            Assert.Empty(result.NewUnboundFields);
            Assert.Equal(boundBefore, result.RemappedCount);
            Assert.True(result.QuickValidation.Passed,
                "expected bindings to still resolve after a namespace-only change; errors: " + string.Join(" | ",
                    result.QuickValidation.Findings.Where(f => f.Severity == FindingSeverity.Error).Select(f => f.Message)));

            using var reopened = WordprocessingDocument.Open(path, false);
            var after = SchemaProvider.FromLayout(reopened);
            Assert.Equal("9999", after.Report.ReportId);
            Assert.Equal(result.StoreItemId, after.Report.StoreItemId, StringComparer.OrdinalIgnoreCase);

            var quick = LayoutValidator.Quick(reopened);
            Assert.Equal(0, quick.ErrorCount);

            // Raw-XML spot check across every part: prefixMappings/tag moved to the new identity, the old
            // one is gone (safe substring check - "1306" vs "9999" never collide as substrings).
            var main = reopened.MainDocumentPart!;
            var allXml = main.Document!.OuterXml
                + string.Concat(main.HeaderParts.Select(h => h.Header?.OuterXml ?? string.Empty))
                + string.Concat(main.FooterParts.Select(f => f.Footer?.OuterXml ?? string.Empty));

            Assert.Contains("Standard_Sales_Invoice/9999/", allXml, StringComparison.Ordinal);
            Assert.DoesNotContain("Standard_Sales_Invoice/1306/", allXml, StringComparison.Ordinal);
            Assert.Contains("#Nav: Standard_Sales_Invoice/9999", allXml, StringComparison.Ordinal);

            // The XPath element-name steps themselves are unchanged - "remap where element names match".
            var inventory = LayoutReader.Read(reopened);
            Assert.Contains(inventory.Controls, c => c.XPath == "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:" + FieldColumn + "[1]");
        }
        finally
        {
            File.Delete(path);
            if (schemaPath is not null)
            {
                File.Delete(schemaPath);
            }
        }
    }

    // ---- error paths ----

    [Fact]
    public void Refresh_missing_newSchemaSource_throws_FileNotFoundException_with_its_path()
    {
        // Note: this only proves the exception itself (type + FileName). The stronger "the layout file on
        // disk is byte-identical after a failed refresh" guarantee is a property of the refresh_xml_part
        // MCP TOOL's copy-then-swap safety net (McpHostToolTests), not of calling Refresh directly against
        // an already-open, already-partially-read WordprocessingDocument - merely opening a package
        // editable and reading a part's raw stream (which SchemaProvider/LayoutReader both do before this
        // exception is ever thrown) can itself perturb the underlying zip container on Dispose regardless
        // of whether anything was actually Save()'d (see ToolGuards.GuardMutate's own remarks on exactly
        // this OpenXml SDK behavior).
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var ex = Assert.Throws<FileNotFoundException>(
                () => LayoutRefresher.Refresh(doc, "Z:\\does-not-exist.xml"));
            Assert.Equal("Z:\\does-not-exist.xml", ex.FileName);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Refresh_non_BC_newSchemaSource_throws_InvalidDataException()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        var badXmlPath = Path.Combine(Path.GetTempPath(), $"bcwl-refresher-badschema-{Guid.NewGuid():N}.xml");
        try
        {
            File.WriteAllText(badXmlPath, "<SomeOtherRoot xmlns=\"urn:not-bc\"><Foo/></SomeOtherRoot>");

            using var doc = WordprocessingDocument.Open(path, true);
            Assert.Throws<InvalidDataException>(() => LayoutRefresher.Refresh(doc, badXmlPath));
        }
        finally
        {
            File.Delete(path);
            if (File.Exists(badXmlPath))
            {
                File.Delete(badXmlPath);
            }
        }
    }
}
