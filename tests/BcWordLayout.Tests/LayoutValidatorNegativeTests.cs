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

    // A repeater over a NESTED data item (needs SyntheticLayout.DatasetWithLineXml). Tests that only need
    // "some well-formed repeater" use this rather than HeaderXPath: a block-level repeater over the
    // top-level item is itself a finding (top-level-repeater-group) and would add an unrelated error.
    private const string LineXPath = Root + "/ns0:Header[1]/ns0:Line";

    // A top-level repeater's binding the way BC writes it: the repeated item's own step carries no index.
    private const string TopLevelHeaderXPath = Root + "/ns0:Header";

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
            Assert.DoesNotContain(result.Findings, f => f.Check == "xpath-resolves");
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
        var bodyRepeater = SyntheticLayout.ProperRepeater(LineXPath, SyntheticLayout.GoodItemId);
        var headerRepeater = SyntheticLayout.ProperRepeater(LineXPath, SyntheticLayout.GoodItemId);
        var path = SyntheticLayout.CreateWithHeader(
            bodyFragments: bodyRepeater, headerFragments: headerRepeater,
            datasetXml: SyntheticLayout.DatasetWithLineXml);

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
            Assert.DoesNotContain(result.Findings, f => f.Check == "repeater-shape");
            Assert.DoesNotContain(result.Findings, f => f.Check == "xpath-resolves");
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

            Assert.DoesNotContain(result.Findings, f => f.Check == "xpath-resolves");
            Assert.DoesNotContain(result.Findings, f => f.Check == "store-item-id");
            Assert.DoesNotContain(result.Findings, f => f.Check == "repeater-shape");
            Assert.DoesNotContain(result.Findings, f => f.Check == "single-bc-part");
            Assert.DoesNotContain(result.Findings, f => f.Check == "binding-namespace");
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

            Assert.DoesNotContain(result.Findings, f => f.Check == "xpath-resolves");
            Assert.DoesNotContain(result.Findings, f => f.Check == "store-item-id");
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

            Assert.DoesNotContain(result.Findings, f => f.Check == "binding-namespace");
            Assert.DoesNotContain(result.Findings, f => f.Check == "binding-prefix-mapping");
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- binding-prefix-mapping: the binding's own declarations must resolve its XPath (GitHub issue #79) ----

    [Theory]
    // The wrong-namespace pitfall SOLUTION-DESIGN §2 calls out: dynamics-365, singular "report".
    [InlineData("xmlns:ns0='urn:microsoft-dynamics-365/report/TestReport/50000/'", "urn:microsoft-dynamics-365/report/TestReport/50000/")]
    // nav, but singular "report".
    [InlineData("xmlns:ns0='urn:microsoft-dynamics-nav/report/TestReport/50000/'", "urn:microsoft-dynamics-nav/report/TestReport/50000/")]
    // The right namespace on the wrong prefix: the XPath uses ns0.
    [InlineData("xmlns:ns1='urn:microsoft-dynamics-nav/reports/TestReport/50000/'", "'ns0:'")]
    // A bare value naming something that is not this layout's namespace.
    [InlineData("urn:microsoft-dynamics-365/report/TestReport/50000/", "bare value")]
    public void Prefix_mapping_that_cannot_resolve_the_xpath_is_an_error_but_a_matching_sibling_is_not(
        string brokenMappings, string messageMustMention)
    {
        // Every one of these shapes was measured unmapped in Word (IsMapped=False) on a copy of JobQuote.docx
        // while quick validation passed it: binding-namespace only sees a namespace that LOOKS like a BC one,
        // and xpath-resolves strips prefixes. The sibling uses the form Word writes and must stay silent.
        var body =
            SyntheticLayout.BoundFieldWithPrefixMappings(
                ValidFieldXPath, SyntheticLayout.GoodItemId, $"xmlns:ns0='{SyntheticLayout.DatasetNamespace}'") +
            SyntheticLayout.BoundFieldWithPrefixMappings(ValidFieldXPath, SyntheticLayout.GoodItemId, brokenMappings);
        var path = SyntheticLayout.Create(body);

        try
        {
            var result = LayoutValidator.Quick(path);

            var finding = Assert.Single(result.Findings, f => f.Check == "binding-prefix-mapping");
            Assert.Equal(FindingSeverity.Error, finding.Severity);
            Assert.Contains(messageMustMention, finding.Message, StringComparison.Ordinal);
            Assert.Contains(SyntheticLayout.DatasetNamespace, finding.Message, StringComparison.Ordinal);
            Assert.False(result.Passed);

            // None of these is binding-namespace's business, and the xpath itself still resolves.
            Assert.DoesNotContain(result.Findings, f => f.Check == "binding-namespace");
            Assert.DoesNotContain(result.Findings, f => f.Check == "xpath-resolves");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    // The form Word and SdtFactory write, with and without the trailing space some captures carry.
    [InlineData("xmlns:ns0='urn:microsoft-dynamics-nav/reports/TestReport/50000/'")]
    [InlineData("xmlns:ns0='urn:microsoft-dynamics-nav/reports/TestReport/50000/' ")]
    // An extra declaration the XPath never uses is harmless.
    [InlineData("xmlns:ns0='urn:microsoft-dynamics-nav/reports/TestReport/50000/' xmlns:ns1='http://example.com/other'")]
    // A BARE uri naming the layout's own namespace: stock layouts ship it and BC prints them, so it is
    // deliberately not reported (see the check's own remarks).
    [InlineData("urn:microsoft-dynamics-nav/reports/TestReport/50000/")]
    public void Prefix_mapping_shapes_that_name_the_layouts_namespace_are_not_flagged(string mappings)
    {
        var path = SyntheticLayout.Create(
            SyntheticLayout.BoundFieldWithPrefixMappings(ValidFieldXPath, SyntheticLayout.GoodItemId, mappings));

        try
        {
            Assert.DoesNotContain(LayoutValidator.Quick(path).Findings, f => f.Check == "binding-prefix-mapping");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Foreign_bc_namespace_is_reported_once_by_binding_namespace_not_again_by_the_prefix_check(bool bareUri)
    {
        // binding-namespace already reports a foreign BC namespace, with the refresh_xml_part repair hint that
        // applies to exactly that case; a second finding for the same wrong namespace would be noise.
        var path = SyntheticLayout.Create(
            SyntheticLayout.BoundFieldInNamespace(
                ValidFieldXPath, SyntheticLayout.GoodItemId, SyntheticLayout.ForeignNamespace, bareUri));

        try
        {
            var result = LayoutValidator.Quick(path);

            Assert.Single(result.Findings, f => f.Check == "binding-namespace");
            Assert.DoesNotContain(result.Findings, f => f.Check == "binding-prefix-mapping");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Undeclared_prefix_is_reported_even_when_binding_namespace_also_fires()
    {
        // Two separate defects on one binding: a foreign BC namespace on ns0 (binding-namespace) and an XPath
        // step on ns1 that nothing declares. refresh_xml_part repairs only the first, so the second must not be
        // hidden behind it.
        var xpath = Root + "/ns0:Header[1]/ns1:CompanyName[1]";
        var path = SyntheticLayout.Create(
            SyntheticLayout.BoundFieldWithPrefixMappings(
                xpath, SyntheticLayout.GoodItemId, $"xmlns:ns0='{SyntheticLayout.ForeignNamespace}'"));

        try
        {
            var result = LayoutValidator.Quick(path);

            Assert.Single(result.Findings, f => f.Check == "binding-namespace");
            var finding = Assert.Single(result.Findings, f => f.Check == "binding-prefix-mapping");
            Assert.Contains("'ns1:'", finding.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Binding_into_another_data_store_is_out_of_the_prefix_checks_scope()
    {
        // A document-property control binds Word's core-properties part, with prefixMappings that name no BC
        // namespace at all. Scoping by the XPath's root step (not NavWordReportXmlPart here) keeps it out.
        const string corePropsMappings =
            "xmlns:ns0='http://purl.org/dc/elements/1.1/' "
            + "xmlns:ns1='http://schemas.openxmlformats.org/package/2006/metadata/core-properties' ";
        var path = SyntheticLayout.Create(
            SyntheticLayout.BoundFieldWithPrefixMappings(
                "/ns1:coreProperties[1]/ns0:title[1]", "{6C3C8BC8-F283-45AE-878A-BAB7291924A1}", corePropsMappings));

        try
        {
            Assert.DoesNotContain(LayoutValidator.Quick(path).Findings, f => f.Check == "binding-prefix-mapping");
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

            Assert.DoesNotContain(result.Findings, f => f.Check == "store-item-id");

            // A warning must not fail the layout: BC re-attaches the store item on upload.
            Assert.True(result.Passed);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- document-content-type: a .docx whose main part is typed as something Word will not open as one (#77) ----

    [Theory]
    [InlineData(DocumentFormat.OpenXml.WordprocessingDocumentType.Template, "template.main+xml")]
    [InlineData(DocumentFormat.OpenXml.WordprocessingDocumentType.MacroEnabledTemplate, "template.macroEnabledTemplate.main+xml")]
    [InlineData(DocumentFormat.OpenXml.WordprocessingDocumentType.MacroEnabledDocument, "document.macroEnabled.main+xml")]
    public void Main_part_not_typed_as_a_document_is_flagged_but_a_document_typed_sibling_is_not(
        DocumentFormat.OpenXml.WordprocessingDocumentType documentType, string contentTypeFragment)
    {
        // A template (or macro-enabled package) renamed to .docx: OpenXmlValidator and every binding check
        // pass it, and Word refuses to open it. The fixtures carry no BC part (single-bc-part fires on both
        // alike), so the pairing isolates the content type as the only difference between them.
        var renamed = SyntheticLayout.CreateWordPackage(documentType, ".docx");
        var sibling = SyntheticLayout.CreateWordPackage(DocumentFormat.OpenXml.WordprocessingDocumentType.Document, ".docx");

        try
        {
            var findings = LayoutValidator.Quick(renamed).Findings.Where(f => f.Check == "document-content-type").ToList();
            var finding = Assert.Single(findings);
            Assert.Equal(FindingSeverity.Error, finding.Severity);
            Assert.Contains(contentTypeFragment, finding.Message, StringComparison.OrdinalIgnoreCase);

            Assert.DoesNotContain(LayoutValidator.Quick(sibling).Findings, f => f.Check == "document-content-type");
        }
        finally
        {
            File.Delete(renamed);
            File.Delete(sibling);
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

    // ---- character-style-resolves: a w:rStyle naming no character style the layout defines (issue #15) ----

    [Fact]
    public void Dangling_or_paragraph_rStyle_references_are_flagged_as_warnings_but_a_defined_character_style_is_not()
    {
        // DefaultStylesScaffold defines the character style DefaultParagraphFont and the paragraph style Normal.
        var body =
            "<w:p><w:r><w:rPr><w:rStyle w:val=\"DefaultParagraphFont\"/></w:rPr><w:t>fine</w:t></w:r></w:p>"
            + "<w:p><w:r><w:rPr><w:rStyle w:val=\"NotDefinedAnywhere\"/></w:rPr><w:t>dangling</w:t></w:r>"
            + "<w:r><w:rPr><w:rStyle w:val=\"NotDefinedAnywhere\"/></w:rPr><w:t>again</w:t></w:r></w:p>"
            + "<w:p><w:r><w:rPr><w:rStyle w:val=\"Normal\"/></w:rPr><w:t>paragraph style</w:t></w:r></w:p>";
        var path = SyntheticLayout.Create(body);

        try
        {
            using (var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(path, true))
            {
                Assert.True(DefaultStylesScaffold.EnsureDefaultStyles(doc.MainDocumentPart!));
            }

            var result = LayoutValidator.Quick(path);

            var styleFindings = result.Findings.Where(f => f.Check == "character-style-resolves").ToList();
            Assert.Equal(2, styleFindings.Count);
            Assert.All(styleFindings, f => Assert.Equal(FindingSeverity.Warning, f.Severity));
            Assert.Contains(styleFindings, f => f.Message.StartsWith(
                "2 run(s) reference character style 'NotDefinedAnywhere' via w:rStyle, but the layout's styles part does not define it",
                StringComparison.Ordinal));
            Assert.Contains(styleFindings, f => f.Message.Contains(
                "'Normal' via w:rStyle, but the layout defines it as a paragraph style, not a character style",
                StringComparison.Ordinal));
            Assert.DoesNotContain(styleFindings, f => f.Message.Contains("DefaultParagraphFont", StringComparison.Ordinal));
            Assert.True(result.Passed);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void No_corpus_layout_trips_character_style_resolves()
    {
        foreach (var file in Directory.GetFiles(Path.Combine(AppContext.BaseDirectory, "corpus"), "*.docx"))
        {
            var result = LayoutValidator.Quick(file);
            Assert.DoesNotContain(result.Findings, f => f.Check == "character-style-resolves");
        }
    }

    [Fact]
    public void Downgraded_repeater_is_flagged_but_a_real_repeater_and_an_ordinary_field_are_not()
    {
        // All three controls in one document, so the check has to discriminate rather than fire on anything
        // that looks vaguely repeater-ish: what Word's compatibility downgrade leaves of a repeater (alias
        // naming a DATA ITEM, no repeatingSection, no binding), a genuine repeater carrying the same alias,
        // and an ordinary bound field whose alias names a LEAF COLUMN (GitHub issue #54).
        var body =
            SyntheticLayout.DowngradedRepeater("/Header") +
            SyntheticLayout.ProperRepeaterWithAlias(HeaderXPath, SyntheticLayout.GoodItemId, "/Header") +
            SyntheticLayout.AliasedBoundField(
                ValidFieldXPath, SyntheticLayout.GoodItemId, "/Header/CompanyName");
        var path = SyntheticLayout.Create(body);

        try
        {
            var result = LayoutValidator.Quick(path);

            var finding = Assert.Single(result.Findings, f => f.Check == "repeater-downgraded");
            Assert.Equal(FindingSeverity.Error, finding.Severity);
            Assert.Contains("/Header", finding.Message);
            Assert.Contains("insert_repeater_table", finding.Message);
            Assert.False(result.Passed);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Word_placeholder_character_scale_is_a_warning_while_other_structural_errors_stay_errors()
    {
        // w:w=0 is what Word writes into a content control's placeholder run, so it appears in any layout
        // that has been opened and saved in Word; ST_TextScale requires >= 1, so OpenXmlValidator reports
        // it. It renders fine in Word and BC and no tool can edit run properties, so failing the layout on
        // it left the caller with an unactionable error (GitHub issue #54). The orphaned
        // repeatingSectionItem alongside it is a REAL structural defect and must keep Error severity - the
        // tolerance is one named shape, not a softening of the gate.
        var body =
            SyntheticLayout.ControlWithCharacterScale(0) +
            SyntheticLayout.OrphanRepeaterItem();
        var path = SyntheticLayout.Create(body);

        try
        {
            var result = LayoutValidator.Quick(path);

            var scale = Assert.Single(
                result.Findings,
                f => f.Check == "openxml-structure" && f.Severity == FindingSeverity.Warning);
            Assert.Contains("Microsoft Word itself writes", scale.Message);

            Assert.Contains(result.Findings, f => f.Check == "repeater-shape");
            Assert.False(result.Passed);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_normal_character_scale_produces_no_structural_finding_at_all()
    {
        // The valid sibling: the same control with a legal w:w value must not be reported in either
        // severity - the tolerance above must not have turned the check into a no-op.
        var path = SyntheticLayout.Create(SyntheticLayout.ControlWithCharacterScale(100));

        try
        {
            var result = LayoutValidator.Quick(path);

            Assert.DoesNotContain(result.Findings, f => f.Check == "openxml-structure");
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
            SyntheticLayout.ProperRepeater(LineXPath, SyntheticLayout.GoodItemId),
            datasetXml: SyntheticLayout.DatasetWithLineXml);

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
            SyntheticLayout.ProperRepeater(LineXPath, SyntheticLayout.GoodItemId),
            datasetXml: SyntheticLayout.DatasetWithLineXml);

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

    // ---- top-level-repeater-group: a block-level repeater over the TOP-LEVEL data item without a w:group
    // around its item renders an EMPTY document in BC (report 1306, verified 2026-10-13). Paired with the
    // three shapes that must NOT trip it: the same repeater wrapped in a w:group (StandardStatement's
    // shape), a block-level repeater over a NESTED item, and a ROW-level repeater over the top-level item
    // (the list-report shape insert_repeater_table builds). ----

    [Fact]
    public void Block_repeater_over_the_top_level_item_without_a_group_is_flagged_but_its_valid_siblings_are_not()
    {
        var body =
            SyntheticLayout.ProperRepeater(TopLevelHeaderXPath, SyntheticLayout.GoodItemId) +
            SyntheticLayout.RepeaterWithField(
                LineXPath, LineXPath + "[1]/ns0:ItemNo[1]", SyntheticLayout.GoodItemId,
                repeaterId: 200, itemId: 201) +
            SyntheticLayout.RepeaterTable(
                TopLevelHeaderXPath, HeaderXPath + "/ns0:CompanyName[1]", SyntheticLayout.GoodItemId);
        var path = SyntheticLayout.Create(body, datasetXml: SyntheticLayout.DatasetWithLineXml);

        try
        {
            var result = LayoutValidator.Quick(path);

            var finding = Assert.Single(result.Findings, f => f.Check == "top-level-repeater-group");
            Assert.Equal(FindingSeverity.Error, finding.Severity);
            Assert.Contains("'Header'", finding.Message);
            Assert.Contains("w:group", finding.Message);
            Assert.Contains("Header[1]", finding.Message);
            Assert.NotNull(finding.Location);
            Assert.Contains("document.xml", finding.Location);
            Assert.False(result.Passed);

            // The defect is this check's alone: the repeater is otherwise well formed.
            Assert.DoesNotContain(result.Findings, f => f.Check == "repeater-shape");
            Assert.DoesNotContain(result.Findings, f => f.Check == "xpath-resolves");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Block_repeater_over_the_top_level_item_inside_a_group_is_not_flagged()
    {
        // The identical repeater with the one sdt that made BC print it: repeatingSection > w:group >
        // repeatingSectionItem. Also proves repeater-shape still accepts the group in between.
        var path = SyntheticLayout.Create(
            SyntheticLayout.GroupedRepeater(TopLevelHeaderXPath, SyntheticLayout.GoodItemId));

        try
        {
            var result = LayoutValidator.Quick(path);

            Assert.DoesNotContain(result.Findings, f => f.Check == "top-level-repeater-group");
            Assert.DoesNotContain(result.Findings, f => f.Check == "repeater-shape");
            Assert.DoesNotContain(result.Findings, f => f.Check == "openxml-structure");
            Assert.True(result.Passed,
                "expected pass; errors: "
                + string.Join(" | ", result.Findings.Where(f => f.Severity == FindingSeverity.Error).Select(f => f.Message)));
        }
        finally
        {
            File.Delete(path);
        }
    }
}
