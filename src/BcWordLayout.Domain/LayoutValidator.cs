using BcWordLayout.Domain.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Domain;

/// <summary>
/// Quick-level structural and binding validation for a BC Word layout. The <c>full</c> (dry-run merge)
/// level lives in <c>BcWordLayout.Merge.FullValidator</c> instead of here, because it exercises
/// <c>MergeEngine</c> and Domain must not take a dependency on the Merge project (Merge depends on Domain).
/// </summary>
public static class LayoutValidator
{
    /// <summary>Runs the quick checks against a layout and returns findings plus an overall pass/fail.</summary>
    public static ValidationResult Quick(string docxPath)
    {
        if (!File.Exists(docxPath))
        {
            throw new FileNotFoundException("The layout file does not exist.", docxPath);
        }

        using var doc = WordprocessingDocument.Open(docxPath, false);
        return Quick(doc);
    }

    /// <summary>Runs the quick checks against an already-open layout package (lets callers share one handle).</summary>
    public static ValidationResult Quick(WordprocessingDocument doc)
    {
        var findings = new List<ValidationFinding>();
        var main = doc.MainDocumentPart
            ?? throw new InvalidDataException("Layout has no main document part.");

        CheckOpenXmlStructure(doc, findings);
        var bcPartItemId = CheckSingleBcPart(main, findings);
        DatasetTree? schema = TryBuildSchema(main, findings);

        var inventory = LayoutReader.Read(doc);
        CheckStoreItemIds(inventory, bcPartItemId, findings);
        CheckBindingNamespaces(inventory, schema, findings);
        CheckXPathsResolve(inventory, schema, findings);
        CheckRepeaterShape(main, findings);
        CheckRepeaterNotInHeaderOrFooter(inventory, findings);
        CheckAttachedTemplate(main, findings);
        CheckTableStylesResolve(main, findings);

        return new ValidationResult { Level = "quick", Findings = findings };
    }

    // 1. OpenXML structural validity.
    private static void CheckOpenXmlStructure(WordprocessingDocument doc, List<ValidationFinding> findings)
    {
        var validator = new OpenXmlValidator(FileFormatVersions.Office2021);
        foreach (var error in validator.Validate(doc))
        {
            findings.Add(new ValidationFinding
            {
                Check = "openxml-structure",
                Severity = FindingSeverity.Error,
                Message = error.Description,
                Location = error.Path?.XPath,
            });
        }
    }

    // 2. Exactly one custom XML part in the BC namespace, and that part carries a DataStoreItem for its
    //    bindings to name. Returns the BC part's item ID (or null when there is none to return).
    private static string? CheckSingleBcPart(MainDocumentPart main, List<ValidationFinding> findings)
    {
        var bcParts = SchemaProvider.FindBcParts(main);

        if (bcParts.Count == 0)
        {
            findings.Add(new ValidationFinding
            {
                Check = "single-bc-part",
                Severity = FindingSeverity.Error,
                Message = $"No BC dataset custom XML part found (expected exactly one with namespace "
                          + $"starting '{OoxmlNames.BcNamespacePrefix}').",
            });
            return null;
        }

        if (bcParts.Count > 1)
        {
            findings.Add(new ValidationFinding
            {
                Check = "single-bc-part",
                Severity = FindingSeverity.Error,
                Message = $"Found {bcParts.Count} BC dataset custom XML parts; exactly one is expected.",
            });
        }

        var itemId = bcParts[0].CustomXmlPropertiesPart?.DataStoreItem?.ItemId?.Value;
        if (itemId is null)
        {
            // The BC part exists but has no CustomXmlPropertiesPart/DataStoreItem, so there is no item ID for
            // CheckStoreItemIds to compare bindings against. Reported here rather than passed over silently:
            // without it, a layout whose every binding names a storeItemID absent from its own package
            // validated as fully passing (PaymentPracticeByPeriod.docx, all 25 bindings). A warning rather
            // than an error because real base-app layouts genuinely ship this way (StandardStatement.docx
            // too) and BC re-attaches the store item itself on upload — the layout is not broken, its
            // bindings simply cannot be verified from the package alone.
            findings.Add(new ValidationFinding
            {
                Check = "single-bc-part",
                Severity = FindingSeverity.Warning,
                Message = "The BC dataset custom XML part has no itemProps/DataStoreItem, so it declares no "
                          + "item ID; binding storeItemIDs cannot be verified against it. Business Central "
                          + "re-attaches the store item on upload, so this is not itself a defect.",
            });
        }

        return itemId;
    }

    private static DatasetTree? TryBuildSchema(MainDocumentPart main, List<ValidationFinding> findings)
    {
        try
        {
            return SchemaProvider.FromMainPart(main);
        }
        catch (Exception ex)
        {
            findings.Add(new ValidationFinding
            {
                Check = "schema-parse",
                Severity = FindingSeverity.Error,
                Message = $"Could not parse the BC dataset schema: {ex.Message}",
            });
            return null;
        }
    }

    // 3. Every binding's storeItemID matches the BC part's actual item ID.
    private static void CheckStoreItemIds(LayoutInventory inventory, string? bcPartItemId, List<ValidationFinding> findings)
    {
        if (bcPartItemId is null)
        {
            // Nothing to compare against. Both reasons are already reported by CheckSingleBcPart — no BC
            // part at all (error), or a BC part with no DataStoreItem to hold an item ID (warning).
            return;
        }

        foreach (var control in inventory.Controls)
        {
            if (control.StoreItemId is null)
            {
                continue;
            }

            if (!GuidEquals(control.StoreItemId, bcPartItemId))
            {
                findings.Add(new ValidationFinding
                {
                    Check = "store-item-id",
                    Severity = FindingSeverity.Error,
                    Message = $"Binding storeItemID '{control.StoreItemId}' does not match the BC part item ID "
                              + $"'{bcPartItemId}'.",
                    Location = $"{control.Part}: {control.Alias ?? control.XPath}",
                });
            }
        }
    }

    // 3b. Every binding's prefixMappings names the layout's OWN BC dataset namespace.
    //
    //     This is the check that actually catches an orphaned binding, and neither of its neighbours can
    //     stand in for it. CheckStoreItemIds compares GUIDs and goes quiet whenever the BC part carries no
    //     DataStoreItem; CheckXPathsResolve compares element-NAME steps with prefixes and indices stripped,
    //     so a binding pointing at a different report resolves happily as long as the two datasets happen to
    //     share a path shape — which sibling reports and successive versions of the same report usually do.
    //     PaymentPracticeByPeriod.docx is the proof: 20 of its 25 bindings name the report's superseded
    //     namespace Payment_Practice/590 while the embedded part is Payment_Practice/685, and both other
    //     checks pass it.
    //
    //     ERROR, settled by sandbox evidence (2026-08-02, GitHub issue #1). BC validates every binding's
    //     prefixMappings against the target report's CURRENT namespace at upload and rejects the layout
    //     with one InvalidPrefixMapping per offending binding. Three uploads of PaymentPracticeByPeriod
    //     against its own report 685 proved it: rejected as captured (20 bindings on the superseded 590);
    //     rejected even with the WHOLE package internally consistent on 590 (so no re-pointing and no
    //     structural matching on the upload path — bindings must name the report's current namespace);
    //     accepted the moment every binding was re-pointed to 685, which clears the file of any other
    //     defect. Stock base-app layouts do ship mismatched and customers print those reports — but the
    //     embedded copies never pass through upload validation, so "BC tolerates it" was only ever true
    //     of a path this tool's callers cannot use. (Whether the embedded copies RENDER those controls
    //     with data or silently blank remains unobserved; nothing here rests on it.) In Word's own model
    //     the mismatch is fatal regardless: w:storeItemID selects the data-store part and the
    //     prefixMappings URI must match that part's namespace for the XPath to match any node.
    //
    //     Witness hygiene: SubcontractorDispatchList (38 such bindings) was once cited here as a second
    //     stock witness, but it carries the same contamination signature as the long-struck
    //     QuantityExplosionofBOM — its part claims Subcontractor_Dispatch_List/99000789 while its bindings
    //     name 50000-range namespaces (incl. FS_YSC_SubcontractorDispatch/50030), i.e. a customized-tenant
    //     capture, not a base-app one. It stays in the corpus for merge-depth coverage only (Corpus.cs).
    private static void CheckBindingNamespaces(LayoutInventory inventory, DatasetTree? schema, List<ValidationFinding> findings)
    {
        if (schema is null)
        {
            return; // Already reported by TryBuildSchema.
        }

        var expected = schema.Report.Namespace;

        foreach (var control in inventory.Controls)
        {
            // Null covers both an unbound control and a binding whose prefixMappings names no BC namespace
            // at all; neither is this check's business.
            if (control.BindingNamespace is null
                || string.Equals(control.BindingNamespace, expected, StringComparison.Ordinal))
            {
                continue;
            }

            findings.Add(new ValidationFinding
            {
                Check = "binding-namespace",
                Severity = FindingSeverity.Error,
                Message = $"Binding names dataset namespace '{control.BindingNamespace}', but this layout's "
                          + $"BC part declares '{expected}'. Business Central rejects the layout at upload "
                          + "with InvalidPrefixMapping for every such binding (sandbox-verified 2026-08-02), "
                          + "and in Word's model the control would render blank. To repair every such binding "
                          + "in one call, run refresh_xml_part against this report's current schema - it "
                          + "re-points every BC-namespaced binding at the layout's own namespace, which is "
                          + "the only state BC's upload validation accepts. That fixes the NAMESPACE only: if "
                          + "the binding's XPath also names a column this report does not have, refresh "
                          + "reports it as an orphan and rebuilding the control (remove_control then "
                          + "insert_field/insert_label) is still the fix.",
                Location = $"{control.Part}: {control.Alias ?? control.XPath}",
            });
        }
    }

    // 4. Every binding's XPath resolves against the parsed schema tree (segment-by-segment).
    private static void CheckXPathsResolve(LayoutInventory inventory, DatasetTree? schema, List<ValidationFinding> findings)
    {
        if (schema is null)
        {
            return; // Already reported by TryBuildSchema.
        }

        foreach (var control in inventory.Controls)
        {
            if (control.XPath is null)
            {
                continue;
            }

            if (!Resolves(control.XPath, schema, out var failedSegment))
            {
                findings.Add(new ValidationFinding
                {
                    Check = "xpath-resolves",
                    Severity = FindingSeverity.Error,
                    Message = $"Binding XPath does not resolve against the schema: segment '{failedSegment}' "
                              + "was not found in the dataset hierarchy.",
                    Location = $"{control.Part}: {control.XPath}",
                });
            }
        }
    }

    /// <summary>
    /// Structural segment-by-segment resolution. Indices and namespace prefixes are stripped. Intermediate
    /// segments must be data items; the final segment may be a data item or a leaf column.
    /// </summary>
    internal static bool Resolves(string xpath, DatasetTree schema, out string failedSegment)
    {
        failedSegment = "";
        var segments = BindingXPath.Segments(xpath);
        if (segments.Count == 0)
        {
            failedSegment = "(empty)";
            return false;
        }

        if (!string.Equals(segments[0], schema.Root.Name, StringComparison.Ordinal))
        {
            failedSegment = segments[0];
            return false;
        }

        var node = schema.Root;
        for (var i = 1; i < segments.Count; i++)
        {
            var name = segments[i];
            var isLast = i == segments.Count - 1;

            var childItem = node.FindChildItem(name);
            if (childItem is not null)
            {
                node = childItem;
                continue;
            }

            if (isLast && node.FindChildColumn(name) is not null)
            {
                return true;
            }

            failedSegment = name;
            return false;
        }

        return true;
    }

    // 5. Repeater shape: each repeatingSection contains exactly one repeatingSectionItem (its own,
    //    not one belonging to a nested repeater), and every repeatingSectionItem is enclosed by a
    //    repeatingSection. This is a structural/tree check — the schema XPath hierarchy may legitimately
    //    skip intermediate data items that have no repeater control.
    private static void CheckRepeaterShape(MainDocumentPart main, List<ValidationFinding> findings)
    {
        foreach (var (rootElement, partName) in EnumerateContentParts(main))
        {
            foreach (var repeater in rootElement.Descendants<SdtElement>().Where(SdtInspector.IsRepeater))
            {
                var ownItems = repeater.Descendants<SdtElement>()
                    .Count(s => SdtInspector.IsRepeaterItem(s) && SdtInspector.NearestRepeaterAncestor(s) == repeater);

                if (ownItems != 1)
                {
                    findings.Add(new ValidationFinding
                    {
                        Check = "repeater-shape",
                        Severity = FindingSeverity.Error,
                        Message = $"Repeating section must contain exactly one repeatingSectionItem, found {ownItems}.",
                        Location = $"{partName}: {RepeaterXPath(repeater)}",
                    });
                }
            }

            // Orphaned repeatingSectionItem: an item sdt with no enclosing repeatingSection (e.g. a
            // copy-paste error duplicating an item outside its repeater) is never counted above, so it
            // must be flagged in its own right.
            foreach (var item in rootElement.Descendants<SdtElement>().Where(SdtInspector.IsRepeaterItem))
            {
                if (SdtInspector.NearestRepeaterAncestor(item) is null)
                {
                    findings.Add(new ValidationFinding
                    {
                        Check = "repeater-shape",
                        Severity = FindingSeverity.Error,
                        Message = "Found a repeatingSectionItem with no enclosing repeatingSection.",
                        Location = $"{partName}: {SdtDescriptor(item)}",
                    });
                }
            }
        }
    }

    // 6. Repeater LOCATION: a repeating section control (w15:repeatingSection) located in a header/footer
    //    part is outside this tool's v1 supported matrix — insert_repeater_table itself now refuses to
    //    CREATE one there (see LayoutEditor.InsertRepeaterTable), but a hand-authored layout, one edited
    //    outside this tool, or a pre-existing BC layout could still legitimately have one; the design's own
    //    principle is to flag anything outside the matrix instead of silently treating it the same as a
    //    body repeater, since Business Central may not merge/render it reliably. Warning (not Error): it
    //    does not indicate the layout is structurally broken, only that it exercises an unverified path.
    private static void CheckRepeaterNotInHeaderOrFooter(LayoutInventory inventory, List<ValidationFinding> findings)
    {
        foreach (var control in inventory.Controls)
        {
            if (control.Kind != ControlKind.Repeater || control.Part == "document.xml")
            {
                continue;
            }

            findings.Add(new ValidationFinding
            {
                Check = "repeater-in-header-footer",
                Severity = FindingSeverity.Warning,
                Message = "This repeating section is located in a header/footer part, outside this tool's v1 "
                          + "supported matrix (repeaters in headers/footers are deferred - GitHub "
                          + "issue #10); Business Central may not merge/render it reliably.",
                Location = $"{control.Part}: {control.Alias ?? control.XPath}",
            });
        }
    }

    // 7. External attachedTemplate relationship → warning (stale developer template path is common).
    private static void CheckAttachedTemplate(MainDocumentPart main, List<ValidationFinding> findings)
    {
        var visited = new HashSet<OpenXmlPart>();
        foreach (var rel in AllExternalRelationships(main, visited))
        {
            if (rel.RelationshipType.EndsWith("attachedTemplate", StringComparison.Ordinal))
            {
                findings.Add(new ValidationFinding
                {
                    Check = "attached-template",
                    Severity = FindingSeverity.Warning,
                    Message = "Layout has an external attachedTemplate relationship. This often points at a "
                              + "stale developer path and is harmless for BC rendering, but consider removing it.",
                    Location = rel.Uri?.ToString(),
                });
            }
        }
    }

    // 8. Every w:tblStyle names a style the layout's own styles part actually defines. A dangling reference
    //    is not structurally invalid and Word/BC simply ignore it — which is exactly the trap: a caller who
    //    passed insert_repeater_table's tableStyle parameter (or hand-authored a w:tblStyle) believes they
    //    styled the table, and every renderer silently falls back to its own defaults instead (GitHub
    //    issue #3's knock-on finding). WARNING, not an error: the layout renders fine, the reference just
    //    does nothing, and BC's upload validation does not police style references. Layouts this tool creates from
    //    scratch ship a styles part defining TableGrid (see DefaultStylesScaffold), so the documented
    //    tableStyle example resolves; this check covers everything else — typo'd style names, bare
    //    externally-authored layouts, and templates whose styles part lacks the referenced style.
    private static void CheckTableStylesResolve(MainDocumentPart main, List<ValidationFinding> findings)
    {
        // Style ids are case-sensitive in Word's model (w:styleId is an xsd:string) — Ordinal, not
        // OrdinalIgnoreCase, so "tablegrid" vs "TableGrid" is correctly reported as unresolved.
        var definedStyleIds = main.StyleDefinitionsPart?.Styles?
            .Elements<Style>()
            .Select(s => s.StyleId?.Value)
            .Where(id => !string.IsNullOrEmpty(id))
            .ToHashSet(StringComparer.Ordinal);

        foreach (var (rootElement, partName) in EnumerateContentParts(main))
        {
            foreach (var tableStyle in rootElement.Descendants<TableStyle>())
            {
                var styleId = tableStyle.Val?.Value;
                if (string.IsNullOrEmpty(styleId) || (definedStyleIds?.Contains(styleId) ?? false))
                {
                    continue;
                }

                findings.Add(new ValidationFinding
                {
                    Check = "table-style-resolves",
                    Severity = FindingSeverity.Warning,
                    Message = $"A table references style '{styleId}' via w:tblStyle, but "
                              + (definedStyleIds is null
                                  ? "this layout has no styles part at all"
                                  : "the layout's styles part does not define it")
                              + ", so the reference silently does nothing — every renderer falls back to "
                              + "its own defaults. Reference a style the layout defines (a layout created "
                              + "by create_layout defines 'TableGrid'), start from a templatePath that "
                              + "defines the style, or drop the style reference.",
                    Location = partName,
                });
            }
        }
    }

    // ---- helpers ----

    private static IEnumerable<(OpenXmlPartRootElement Root, string PartName)> EnumerateContentParts(MainDocumentPart main) =>
        PartWalker.ContentParts(main);

    private static string? RepeaterXPath(SdtElement repeater)
    {
        var pr = repeater.GetFirstChild<SdtProperties>();
        if (pr is null)
        {
            return null;
        }

        var binding = SdtInspector.FindRepeaterBinding(pr);
        return binding is null ? null : SdtInspector.Attr(binding, "xpath", OoxmlNames.W);
    }

    // Best-effort locator for an sdt without a w15 binding (e.g. a repeatingSectionItem): its w:alias,
    // then its w:tag, then its w:id.
    private static string SdtDescriptor(SdtElement sdt)
    {
        var pr = sdt.GetFirstChild<SdtProperties>();
        if (pr is null)
        {
            return "(sdt)";
        }

        var alias = SdtInspector.ChildVal(pr, "alias", OoxmlNames.W);
        if (!string.IsNullOrEmpty(alias))
        {
            return $"alias '{alias}'";
        }

        var tag = SdtInspector.ChildVal(pr, "tag", OoxmlNames.W);
        if (!string.IsNullOrEmpty(tag))
        {
            return $"tag '{tag}'";
        }

        var id = SdtInspector.ChildVal(pr, "id", OoxmlNames.W);
        return id is null ? "(sdt)" : $"id {id}";
    }

    /// <summary>
    /// <paramref name="depth"/> (0 at the package root's direct parts) enforces
    /// <see cref="ResourceLimits.MaxPartGraphDepth"/> (an uncatchable <see cref="StackOverflowException"/>
    /// would kill the server; a cap fails the call instead): the
    /// <paramref name="visited"/> set already rules out an infinite CYCLE, but a crafted package could still
    /// chain many thousands of otherwise-acyclic parts, making this hand-rolled recursive walk itself
    /// arbitrarily deep.
    /// </summary>
    private static IEnumerable<ExternalRelationship> AllExternalRelationships(OpenXmlPart part, HashSet<OpenXmlPart> visited, int depth = 0)
    {
        if (depth > ResourceLimits.MaxPartGraphDepth)
        {
            throw ResourceLimits.DepthExceeded("Package part graph", ResourceLimits.MaxPartGraphDepth);
        }

        if (!visited.Add(part))
        {
            yield break;
        }

        foreach (var rel in part.ExternalRelationships)
        {
            yield return rel;
        }

        foreach (var child in part.Parts)
        {
            foreach (var rel in AllExternalRelationships(child.OpenXmlPart, visited, depth + 1))
            {
                yield return rel;
            }
        }
    }

    private static bool GuidEquals(string a, string b)
    {
        static string Norm(string s) => s.Trim().Trim('{', '}').ToLowerInvariant();
        return Norm(a) == Norm(b);
    }
}
