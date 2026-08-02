using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;
using BcWordLayout.Merge;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

/// <summary>
/// Two things the 2026-08-01 corpus additions turned from "defensive code nobody had exercised" into
/// pinned behaviour.
/// <para>
/// B39 — <b>the second field-control shape.</b> <see cref="SdtInspector.FindBinding(SdtElement)"/> falls back
/// from <c>w:dataBinding</c> to <c>w15:dataBinding</c> for field/label controls, added on the reasoning that
/// "the fallback costs nothing" with no known instance. It is not hypothetical: 367 such controls were
/// measured across the reviewed layouts, most carrying no <c>&lt;w:text/&gt;</c> marker and many no
/// <c>w:alias</c>/<c>w:tag</c> at all — so the <c>#Nav:</c> alias/tag convention that error messages and hints
/// lean on is simply absent. <c>PaymentPracticeByPeriod.docx</c> uses this shape EXCLUSIVELY (not one legacy
/// field control anywhere) and <c>JobQuote.docx</c> mixes all three variants in one document.
/// </para>
/// <para>
/// B40 — <b>the deepest nesting.</b> <c>SubcontractorDispatchList.docx</c> is a straight FIVE-level repeater
/// chain with every level inside one table, one deeper than <c>StandardStatement.docx</c> and a different
/// shape (pure depth rather than a branching tree that skips data items). Since
/// <see cref="XPathReanchor"/>'s per-level step-drop compounds with depth, this is the case most likely to
/// expose an off-by-one — so the merge is asserted level by level rather than just "it did not throw".
/// </para>
/// </summary>
public class W15BindingAndDeepNestingTests
{
    private static string TempOutput() =>
        Path.Combine(Path.GetTempPath(), $"bcwl-w15deep-{Guid.NewGuid():N}.docx");

    private static List<ValidationErrorInfo> OpenXmlErrors(WordprocessingDocument doc) =>
        new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc).ToList();

    // ---- B39: w15:dataBinding field controls ----

    [Fact]
    public void A_layout_bound_entirely_through_w15_dataBinding_still_reads_as_Field_controls()
    {
        using var doc = WordprocessingDocument.Open(Corpus.Path(Corpus.PaymentPracticeByPeriod), false);
        var inventory = LayoutReader.Read(doc);

        var fields = inventory.Controls.Where(c => c.Kind == ControlKind.Field).ToList();
        Assert.Equal(24, fields.Count);

        // Every one of them is w15-bound: if FindBinding's fallback were removed, all 24 would collapse to
        // Unbound and the layout would read as having no data at all.
        Assert.All(fields, f => Assert.True(f.UsesW15Binding));
        Assert.All(fields, f => Assert.NotNull(f.XPath));

        // And the shape really is the alias-less/tag-less/text-less one - so nothing downstream may require
        // an alias to identify a control.
        Assert.Contains(fields, f => f.Alias is null && f.Tag is null);

        Assert.Single(inventory.Controls.Where(c => c.Kind == ControlKind.Repeater));
    }

    [Fact]
    public void All_three_field_control_shapes_coexist_in_one_layout_and_all_read_as_Field()
    {
        // JobQuote carries legacy w:dataBinding + <w:text/> + alias/tag, w15:dataBinding WITH alias/tag, and
        // w15:dataBinding with neither. All three must classify identically.
        using var doc = WordprocessingDocument.Open(Corpus.Path(Corpus.JobQuote), false);
        var controls = LayoutReader.Read(doc).Controls;

        var bound = controls.Where(c => c.Kind is ControlKind.Field or ControlKind.Label).ToList();
        Assert.Contains(bound, c => !c.UsesW15Binding);                       // legacy
        Assert.Contains(bound, c => c.UsesW15Binding && c.Alias is not null); // w15 + alias
        Assert.Contains(bound, c => c.UsesW15Binding && c.Alias is null);     // w15, alias-less
        Assert.All(bound, c => Assert.NotNull(c.XPath));
    }

    [Fact]
    public void FindBinding_falls_back_to_w15_when_a_field_carries_no_legacy_dataBinding()
    {
        // The unit-level statement of the same fact, so the fallback cannot be deleted as dead code even if
        // the corpus files were ever swapped out.
        const string xpath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:CompanyName[1]";
        var body =
            "<w:sdt><w:sdtPr>"
            + $"<w15:dataBinding w:xpath=\"{xpath}\" w:storeItemID=\"{SyntheticLayout.GoodItemId}\"/>"
            + "</w:sdtPr><w:sdtContent><w:p><w:r><w:t>x</w:t></w:r></w:p></w:sdtContent></w:sdt>";
        var path = SyntheticLayout.Create(body);
        try
        {
            using var doc = WordprocessingDocument.Open(path, false);
            var control = Assert.Single(LayoutReader.Read(doc).Controls);

            Assert.Equal(ControlKind.Field, control.Kind);
            Assert.True(control.UsesW15Binding);
            Assert.Equal(xpath, control.XPath);
            Assert.Equal(SyntheticLayout.GoodItemId, control.StoreItemId);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Merge_fills_w15_bound_fields_with_sample_data()
    {
        // Reading them is not enough: the merge/preview pillar has to fill them too, and it identifies
        // controls through the same FindBinding seam.
        var output = TempOutput();
        try
        {
            var result = MergeEngine.Merge(Corpus.Path(Corpus.PaymentPracticeByPeriod), output);

            Assert.True(result.Stats.FieldsFilled > 0,
                "expected the w15-bound fields to be filled; warnings: "
                + string.Join(" | ", result.Warnings.Select(w => w.Message)));

            using var merged = WordprocessingDocument.Open(output, false);
            Assert.Empty(OpenXmlErrors(merged));
        }
        finally
        {
            File.Delete(output);
        }
    }

    // ---- B40: the five-level repeater chain ----

    [Fact]
    public void The_deepest_corpus_nesting_is_five_levels_in_a_single_table()
    {
        // Guards the premise of the merge test below: if this layout were ever replaced by a shallower one,
        // that test would still pass while silently covering nothing special.
        using var doc = WordprocessingDocument.Open(Corpus.Path(Corpus.SubcontractorDispatchList), false);
        var controls = LayoutReader.Read(doc).Controls;

        var repeaters = controls.Where(c => c.Kind == ControlKind.Repeater).ToList();
        Assert.Equal(5, repeaters.Count);

        // A straight chain: exactly one repeater has no repeater parent, and each of the others has one.
        Assert.Single(repeaters.Where(r => r.ParentRepeater is null));

        static int Depth(LayoutControl c) => c.ParentRepeater is null ? 1 : 1 + Depth(c.ParentRepeater);
        Assert.Equal(5, repeaters.Max(Depth));

        // All five live in the body, in the SAME table - the shape that makes this different from
        // StandardStatement, whose levels branch across several tables. Only non-null indices are compared:
        // the outermost repeater wraps the table's row rather than sitting in a cell, so it reports none.
        Assert.All(repeaters, r => Assert.Equal("document.xml", r.Part));
        var tableIndices = repeaters.Where(r => r.TableIndex is not null).Select(r => r.TableIndex).Distinct().ToList();
        Assert.NotEmpty(tableIndices);
        Assert.Single(tableIndices);
    }

    [Fact]
    public void Merge_expands_the_five_level_chain_and_reanchors_every_level()
    {
        var output = TempOutput();
        try
        {
            var result = MergeEngine.Merge(Corpus.Path(Corpus.SubcontractorDispatchList), output);

            // Every level expanded: 5 repeaters each cloning at least one row.
            Assert.True(result.Stats.RepeatersExpanded >= 5,
                $"expected all 5 levels expanded, got {result.Stats.RepeatersExpanded}");

            using var merged = WordprocessingDocument.Open(output, false);

            // The merged document is structurally valid AND its table grid survived five levels of row
            // cloning - the guard OpenXmlValidator cannot see.
            Assert.Empty(OpenXmlErrors(merged));
            Assert.Empty(TableGridConsistencyGuard.Find(merged));

            // Re-anchoring worked at depth. Asserted on the machine-readable warning Kind and on
            // Stats.Unresolved rather than on message text, so a reworded message cannot quietly weaken it.
            var unresolved = result.Warnings.Where(w => w.Kind == "unresolved-binding").ToList();
            Assert.True(unresolved.Count == 0,
                "unexpected unresolved bindings: " + string.Join(" | ", unresolved.Select(w => w.Message)));
            Assert.Equal(0, result.Stats.Unresolved);
        }
        finally
        {
            File.Delete(output);
        }
    }
}
