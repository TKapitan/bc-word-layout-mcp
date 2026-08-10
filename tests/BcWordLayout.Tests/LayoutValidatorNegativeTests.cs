using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;

namespace BcWordLayout.Tests;

/// <summary>
/// Negative-path coverage for <see cref="LayoutValidator"/> using deliberately malformed synthetic
/// layouts. These prove the checks actually flag broken input — the corpus tests only prove they pass
/// on known-good input. Non-tautology is enforced by pairing each defect with a valid sibling control
/// that must NOT trip the same check.
/// </summary>
public class LayoutValidatorNegativeTests
{
    private const string Root = "/ns0:NavWordReportXmlPart[1]";
    private const string ValidFieldXPath = Root + "/ns0:Header[1]/ns0:CompanyName[1]";
    private const string BogusFieldXPath = Root + "/ns0:Header[1]/ns0:BogusField[1]";
    private const string HeaderXPath = Root + "/ns0:Header[1]";

    [Fact]
    public void XPath_to_nonexistent_field_is_flagged_but_valid_sibling_is_not()
    {
        var body =
            SyntheticLayout.BoundField(ValidFieldXPath, SyntheticLayout.GoodItemId) +
            SyntheticLayout.BoundField(BogusFieldXPath, SyntheticLayout.GoodItemId);
        var path = SyntheticLayout.Create(body);

        try
        {
            var result = LayoutValidator.Quick(path);

            var xpathFindings = result.Findings.Where(f => f.Check == "xpath-resolves").ToList();
            Assert.Single(xpathFindings);
            Assert.Contains("BogusField", xpathFindings[0].Message);
            Assert.Equal(FindingSeverity.Error, xpathFindings[0].Severity);
            Assert.False(result.Passed);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void StoreItemId_mismatch_is_flagged_but_matching_sibling_is_not()
    {
        // Both bindings have valid xpaths, so xpath-resolves stays silent and isolates the storeItemID check.
        var body =
            SyntheticLayout.BoundField(ValidFieldXPath, SyntheticLayout.GoodItemId) +
            SyntheticLayout.BoundField(ValidFieldXPath, SyntheticLayout.WrongItemId);
        var path = SyntheticLayout.Create(body, partItemId: SyntheticLayout.GoodItemId);

        try
        {
            var result = LayoutValidator.Quick(path);

            var storeFindings = result.Findings.Where(f => f.Check == "store-item-id").ToList();
            Assert.Single(storeFindings);
            Assert.Contains(SyntheticLayout.WrongItemId, storeFindings[0].Message);
            Assert.Empty(result.Findings.Where(f => f.Check == "xpath-resolves"));
            Assert.False(result.Passed);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Orphaned_repeatingSectionItem_is_flagged_but_a_proper_repeater_is_not()
    {
        var body =
            SyntheticLayout.ProperRepeater(HeaderXPath, SyntheticLayout.GoodItemId) +
            SyntheticLayout.OrphanRepeaterItem();
        var path = SyntheticLayout.Create(body);

        try
        {
            var result = LayoutValidator.Quick(path);

            var shapeFindings = result.Findings.Where(f => f.Check == "repeater-shape").ToList();
            Assert.Single(shapeFindings);
            Assert.Contains("no enclosing repeatingSection", shapeFindings[0].Message);
            Assert.Equal(FindingSeverity.Error, shapeFindings[0].Severity);
            Assert.False(result.Passed);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- Follow-up review (post-Phase-4.3): a repeater LOCATED in a header/footer part is outside the v1
    // supported matrix - flagged as a Warning (not an Error: the layout isn't structurally broken, it just
    // exercises an unverified path), paired with a body repeater that must NOT trip the same check. ----

    [Fact]
    public void Repeater_located_in_a_header_part_is_flagged_as_a_warning_but_a_body_repeater_is_not()
    {
        var bodyRepeater = SyntheticLayout.ProperRepeater(HeaderXPath, SyntheticLayout.GoodItemId);
        var headerRepeater = SyntheticLayout.ProperRepeater(HeaderXPath, SyntheticLayout.GoodItemId);
        var path = SyntheticLayout.CreateWithHeader(bodyFragments: bodyRepeater, headerFragments: headerRepeater);

        try
        {
            var result = LayoutValidator.Quick(path);

            var locationFindings = result.Findings.Where(f => f.Check == "repeater-in-header-footer").ToList();
            var finding = Assert.Single(locationFindings);
            Assert.Equal(FindingSeverity.Warning, finding.Severity);
            Assert.NotNull(finding.Location);
            Assert.Contains("header", finding.Location, StringComparison.OrdinalIgnoreCase);

            // A Warning must never fail quick validation on its own.
            Assert.True(result.Passed, "expected a Warning-only finding not to fail Passed");
            Assert.Equal(0, result.ErrorCount);

            // No structural/shape/binding errors from either repeater - isolates this new check from the
            // ones covered elsewhere in this file.
            Assert.Empty(result.Findings.Where(f => f.Check == "repeater-shape"));
            Assert.Empty(result.Findings.Where(f => f.Check == "xpath-resolves"));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Well_formed_synthetic_layout_produces_no_binding_or_shape_errors()
    {
        // Guards against the negative tests silently passing due to unrelated always-on errors.
        var body =
            SyntheticLayout.BoundField(ValidFieldXPath, SyntheticLayout.GoodItemId) +
            SyntheticLayout.ProperRepeater(HeaderXPath, SyntheticLayout.GoodItemId);
        var path = SyntheticLayout.Create(body);

        try
        {
            var result = LayoutValidator.Quick(path);

            Assert.Empty(result.Findings.Where(f => f.Check == "xpath-resolves"));
            Assert.Empty(result.Findings.Where(f => f.Check == "store-item-id"));
            Assert.Empty(result.Findings.Where(f => f.Check == "repeater-shape"));
            Assert.Empty(result.Findings.Where(f => f.Check == "single-bc-part"));
            Assert.Empty(result.Findings.Where(f => f.Check == "binding-namespace"));
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- binding-namespace: the check that catches an orphaned binding neither neighbour can see ----

    [Theory]
    [InlineData(false)] // xmlns:ns0='<uri>' — the form Word writes
    [InlineData(true)]  // bare <uri> — the form real base-app layouts also ship (see BoundFieldInNamespace)
    public void Binding_to_a_foreign_dataset_namespace_is_flagged_but_a_matching_sibling_is_not(bool bareUri)
    {
        // Both bindings use a resolvable xpath and the right storeItemID, so xpath-resolves and
        // store-item-id both stay silent and isolate the namespace check. That isolation is the point:
        // these two checks CANNOT catch a foreign namespace, which is why this check exists.
        var body =
            SyntheticLayout.BoundFieldInNamespace(
                ValidFieldXPath, SyntheticLayout.GoodItemId, SyntheticLayout.DatasetNamespace, bareUri) +
            SyntheticLayout.BoundFieldInNamespace(
                ValidFieldXPath, SyntheticLayout.GoodItemId, SyntheticLayout.ForeignNamespace, bareUri);
        var path = SyntheticLayout.Create(body);

        try
        {
            var result = LayoutValidator.Quick(path);

            var nsFindings = result.Findings.Where(f => f.Check == "binding-namespace").ToList();
            Assert.Single(nsFindings);
            Assert.Contains(SyntheticLayout.ForeignNamespace, nsFindings[0].Message);
            Assert.Contains(SyntheticLayout.DatasetNamespace, nsFindings[0].Message);

            // Error since the 2026-08-02 sandbox verification (GitHub issue #1): BC rejects any layout
            // whose binding names a namespace other than the target report's current one at upload
            // (InvalidPrefixMapping), so validate_layout must fail it too.
            Assert.Equal(FindingSeverity.Error, nsFindings[0].Severity);
            Assert.False(result.Passed);

            Assert.Empty(result.Findings.Where(f => f.Check == "xpath-resolves"));
            Assert.Empty(result.Findings.Where(f => f.Check == "store-item-id"));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Binding_with_no_prefixMappings_at_all_is_not_flagged_by_the_namespace_check()
    {
        // A binding may legitimately carry no prefixMappings; there is then no declared namespace to
        // disagree with, and inventing a complaint would fire on every synthetic and hand-built layout.
        var path = SyntheticLayout.Create(
            SyntheticLayout.BoundField(ValidFieldXPath, SyntheticLayout.GoodItemId));

        try
        {
            var result = LayoutValidator.Quick(path);

            Assert.Empty(result.Findings.Where(f => f.Check == "binding-namespace"));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Bc_part_without_a_store_item_id_warns_instead_of_silently_skipping_the_storeItemId_check()
    {
        // The B36 regression. The WrongItemId binding is unverifiable here because the part declares no
        // item ID at all - so store-item-id must stay silent (nothing to compare) while single-bc-part
        // says out loud that it could not be checked. Before the fix, BOTH were silent and a layout whose
        // every binding named an absent part validated as passing with zero findings.
        var path = SyntheticLayout.CreateWithoutStoreItemId(
            SyntheticLayout.BoundField(ValidFieldXPath, SyntheticLayout.WrongItemId));

        try
        {
            var result = LayoutValidator.Quick(path);

            var partFindings = result.Findings.Where(f => f.Check == "single-bc-part").ToList();
            Assert.Single(partFindings);
            Assert.Equal(FindingSeverity.Warning, partFindings[0].Severity);
            Assert.Contains("DataStoreItem", partFindings[0].Message);

            Assert.Empty(result.Findings.Where(f => f.Check == "store-item-id"));

            // A warning must not fail the layout: BC re-attaches the store item on upload.
            Assert.True(result.Passed);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- table-style-resolves: a w:tblStyle naming a style the layout does not define (issue #3) ----

    [Fact]
    public void Dangling_tblStyle_reference_is_flagged_as_a_warning_but_a_resolving_sibling_is_not()
    {
        // Two styled tables in one document: one referencing TableGrid (defined below by the same
        // DefaultStylesScaffold a blank create_layout ships, dogfooded here as the fixture's styles part)
        // and one referencing a style nothing defines. Only the latter may trip the check.
        var body =
            SyntheticLayout.SimpleStyledTable("TableGrid") +
            SyntheticLayout.SimpleStyledTable("NotDefinedAnywhere");
        var path = SyntheticLayout.Create(body);

        try
        {
            using (var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(path, true))
            {
                Assert.True(DefaultStylesScaffold.EnsureDefaultStyles(doc.MainDocumentPart!));
            }

            var result = LayoutValidator.Quick(path);

            var styleFindings = result.Findings.Where(f => f.Check == "table-style-resolves").ToList();
            var finding = Assert.Single(styleFindings);
            Assert.Equal(FindingSeverity.Warning, finding.Severity);
            Assert.Contains("NotDefinedAnywhere", finding.Message);
            Assert.Equal("document.xml", finding.Location);

            // A dangling style reference renders fine (it just silently does nothing), so it must never
            // fail the layout on its own.
            Assert.True(result.Passed);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void TblStyle_reference_in_a_layout_with_no_styles_part_at_all_is_flagged_as_a_warning()
    {
        // The exact pre-fix create_layout shape: a table referencing the documented 'TableGrid' example in
        // a document that ships no styles part whatsoever - the reference cannot resolve regardless of its
        // spelling, and the message must say the styles part itself is missing.
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleStyledTable("TableGrid"));

        try
        {
            var result = LayoutValidator.Quick(path);

            var styleFindings = result.Findings.Where(f => f.Check == "table-style-resolves").ToList();
            var finding = Assert.Single(styleFindings);
            Assert.Equal(FindingSeverity.Warning, finding.Severity);
            Assert.Contains("no styles part", finding.Message);
            Assert.True(result.Passed);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Repeater_in_a_layout_declaring_no_compatibility_mode_is_flagged_as_a_warning()
    {
        // The exact pre-fix create_layout shape: a repeating section in a document that ships no settings
        // part, so Word implies mode 12 - where repeating sections do not exist and an interactive save
        // converts them to plain rich-text controls, dropping the binding (GitHub issue #51).
        var path = SyntheticLayout.Create(
            SyntheticLayout.ProperRepeater(HeaderXPath, SyntheticLayout.GoodItemId));

        try
        {
            var result = LayoutValidator.Quick(path);

            var finding = Assert.Single(result.Findings, f => f.Check == "compatibility-mode");
            Assert.Equal(FindingSeverity.Warning, finding.Severity);
            Assert.Contains("declares none", finding.Message);
            Assert.Contains("File > Info > Convert", finding.Message);

            // The layout itself is correct - BC merges the repeater regardless of compatibility mode. This
            // is a Word-round-trip risk, so it must never fail the layout on its own.
            Assert.True(result.Passed);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Repeater_in_a_layout_declaring_compatibility_mode_15_is_not_flagged()
    {
        // The valid sibling of the test above, through the real scaffold: applying
        // DocumentSettingsScaffold to the same fixture must silence the check - which is simultaneously the
        // proof that what a blank build now ships actually satisfies it.
        var path = SyntheticLayout.Create(
            SyntheticLayout.ProperRepeater(HeaderXPath, SyntheticLayout.GoodItemId));

        try
        {
            using (var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(path, true))
            {
                Assert.True(DocumentSettingsScaffold.EnsureCompatibilityMode(doc.MainDocumentPart!));
            }

            var result = LayoutValidator.Quick(path);

            Assert.DoesNotContain(result.Findings, f => f.Check == "compatibility-mode");
            Assert.True(result.Passed);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Layout_with_no_repeater_is_not_flagged_for_its_compatibility_mode()
    {
        // The second valid sibling, and the one that keeps the check from being indiscriminate: a mode-12
        // layout with no repeating section has nothing to lose to the Compatibility Checker, so warning
        // about it would fire on ordinary field-only layouts for a risk that does not exist there.
        var path = SyntheticLayout.Create(
            SyntheticLayout.BoundField(ValidFieldXPath, SyntheticLayout.GoodItemId));

        try
        {
            var result = LayoutValidator.Quick(path);

            Assert.DoesNotContain(result.Findings, f => f.Check == "compatibility-mode");
            Assert.True(result.Passed);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
