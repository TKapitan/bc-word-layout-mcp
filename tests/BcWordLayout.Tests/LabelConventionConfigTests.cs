using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;
using BcWordLayout.McpHost;
using BcWordLayout.McpHost.Tools;
using BcWordLayout.Merge;

namespace BcWordLayout.Tests;

/// <summary>
/// Covers the CONFIGURABLE part of <see cref="LabelConvention"/> (
/// finding P4): a custom suffix list driving <c>insert_label</c>/<c>insert_field</c>/<c>list_dataset_fields</c>
/// consistently through the real tool surface, and the default-on <c>&lt;Labels&gt;</c> data-item rule at the
/// lower <see cref="LabelConvention.IsLabelPath(System.Collections.Generic.IReadOnlyList{string})"/> level.
/// </summary>
/// <remarks>Joins the label-convention-seam collection: every test here swaps
/// <c>LabelConvention.Current</c> (restored in a <c>finally</c> regardless of outcome) — see
/// <see cref="LabelConventionSeamCollection"/> for why every class whose assertions could be disturbed by
/// that swap must be serialized against it.</remarks>
[Collection("label-convention-seam")]
public class LabelConventionConfigTests
{
    // A minimal schema with one plain field (CompanyName) and one "Caption"-suffixed column
    // (CompanyNameCaption) - the shape the BC-default convention ("Lbl" only) does NOT recognize as a label,
    // but a custom ["Lbl", "Caption"] convention does (mirrors the real InventoryOrderDetails.docx shape's
    // *Caption columns, see LabelConvention's own remarks).
    private const string CaptionColumnDatasetXml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        + "<NavWordReportXmlPart xmlns=\"urn:microsoft-dynamics-nav/reports/TestReport/50000/\">"
        + "<BCReportInformation><CreationDateTime>2026-01-01</CreationDateTime></BCReportInformation>"
        + "<Header><CompanyName>Contoso</CompanyName>"
        + "<CompanyNameCaption>CompanyNameCaption</CompanyNameCaption></Header>"
        + "</NavWordReportXmlPart>";

    [Fact]
    public void Custom_suffix_convention_lets_insert_label_accept_a_Caption_column_and_insert_field_reject_it()
    {
        var previous = LabelConvention.Current;
        LabelConvention.Current = new LabelConvention(new[] { "Lbl", "Caption" });
        var path = SyntheticLayout.Create(string.Empty, datasetXml: CaptionColumnDatasetXml);
        try
        {
            var labelResponse = EditTools.InsertLabel(path, "/Header/CompanyNameCaption", "documentEnd");
            Assert.True(labelResponse.Ok, labelResponse.Error?.Message);
            var editDto = Assert.IsType<EditResultDto>(labelResponse.Data);
            Assert.Equal("Label", editDto.Kind);

            var fieldResponse = EditTools.InsertField(path, "/Header/CompanyNameCaption", "documentEnd");
            Assert.False(fieldResponse.Ok);
            Assert.Equal("invalid_argument", fieldResponse.Error!.Code);
            Assert.Contains("field", fieldResponse.Error.Message);

            var fieldsResponse = ReadTools.ListDatasetFields(path);
            var fieldsDto = Assert.IsType<DatasetFieldsDto>(fieldsResponse.Data);
            var header = fieldsDto.Root.Children.Single(c => c.Name == "Header");
            Assert.True(header.Columns.Single(c => c.Name == "CompanyNameCaption").IsLabel);
            Assert.False(header.Columns.Single(c => c.Name == "CompanyName").IsLabel);
        }
        finally
        {
            LabelConvention.Current = previous;
            File.Delete(path);
        }
    }

    [Fact]
    public void Merge_hints_at_the_labels_data_item_knob_when_the_rule_is_disabled()
    {
        // InventoryOrderDetails.docx is the corpus proof of the <Labels> data-item shape (see
        // LabelConvention's remarks). The default convention classifies it out of the box, so the hint can
        // only fire when a host has explicitly disabled (BCWL_LABELS_DATA_ITEM="-") or retargeted the rule:
        // then the ~27 caption columns are sampled as fields and the preview's column headings degrade to
        // raw numbers/dates. The merge must say so and name the env var that restores the rule - the
        // opt-out would otherwise degrade previews silently.
        var previous = LabelConvention.Current;
        LabelConvention.Current = new LabelConvention(new[] { "Lbl" });
        var output = System.IO.Path.Combine(
            System.IO.Path.GetTempPath(), $"bcwl-labels-hint-{Guid.NewGuid():N}.docx");
        try
        {
            var result = MergeEngine.Merge(Corpus.Path(Corpus.InventoryOrderDetails), output);

            var hint = Assert.Single(result.Warnings, w => w.Kind == "labels-convention-hint");
            Assert.Contains("BCWL_LABELS_DATA_ITEM", hint.Message);
            Assert.Contains("'Labels'", hint.Message);
        }
        finally
        {
            LabelConvention.Current = previous;
            File.Delete(output);
        }
    }

    [Fact]
    public void Merge_emits_no_labels_hint_under_the_default_convention()
    {
        // The default convention's labels-data-item rule already classifies the <Labels> shape, so a
        // stock server merging InventoryOrderDetails.docx has nothing to hint about.
        var output = System.IO.Path.Combine(
            System.IO.Path.GetTempPath(), $"bcwl-labels-hint-{Guid.NewGuid():N}.docx");
        try
        {
            var result = MergeEngine.Merge(Corpus.Path(Corpus.InventoryOrderDetails), output);

            Assert.DoesNotContain(result.Warnings, w => w.Kind == "labels-convention-hint");
        }
        finally
        {
            File.Delete(output);
        }
    }

    [Fact]
    public void Merge_emits_no_labels_hint_for_a_schema_without_a_Labels_data_item()
    {
        var output = System.IO.Path.Combine(
            System.IO.Path.GetTempPath(), $"bcwl-labels-hint-{Guid.NewGuid():N}.docx");
        try
        {
            var result = MergeEngine.Merge(Corpus.Path(Corpus.SalesInvoice), output);

            Assert.DoesNotContain(result.Warnings, w => w.Kind == "labels-convention-hint");
        }
        finally
        {
            File.Delete(output);
        }
    }

    [Fact]
    public void Default_convention_still_rejects_the_same_Caption_column_as_a_label()
    {
        // Companion negative case, pinned WITHOUT any swap (LabelConvention.Current stays at its process
        // default here) - proves the widened acceptance above really came from the custom convention, not
        // from some other change to insert_label's own logic.
        var path = SyntheticLayout.Create(string.Empty, datasetXml: CaptionColumnDatasetXml);
        try
        {
            var response = EditTools.InsertLabel(path, "/Header/CompanyNameCaption", "documentEnd");
            Assert.False(response.Ok);
            Assert.Equal("invalid_argument", response.Error!.Code);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Labels_data_item_rule_classifies_every_direct_column_as_a_label_regardless_of_suffix()
    {
        // Lower-level than the tool-surface tests above: exercises LabelConvention.IsLabelPath directly
        // against an in-memory schema shaped like the one real corpus file that actually has this shape -
        // InventoryOrderDetails.docx's dedicated <Labels> data item (see LabelConvention's own remarks) -
        // rather than the shared corpus file itself, so this test cannot disturb (or be disturbed by) any
        // other class's use of that real file.
        var previous = LabelConvention.Current;
        LabelConvention.Current = new LabelConvention(new[] { "Lbl" }, labelsDataItemName: "Labels");
        try
        {
            var labelsItem = new DataItem { Name = "Labels", Path = "/Labels" };
            var unsuffixed = new DatasetColumn { Name = "DataRetrieved", Path = "/Labels/DataRetrieved" };
            var captionSuffixed = new DatasetColumn { Name = "No_ItemCaption", Path = "/Labels/No_ItemCaption" };
            labelsItem.Columns.Add(unsuffixed);
            labelsItem.Columns.Add(captionSuffixed);

            var header = new DataItem { Name = "Header", Path = "/Header" };
            var plainField = new DatasetColumn { Name = "CustomerAddress1", Path = "/Header/CustomerAddress1" };
            header.Columns.Add(plainField);

            // "DataRetrieved" has NO recognized suffix at all - only the Labels-data-item rule can find it.
            Assert.True(unsuffixed.IsLabel);
            Assert.True(captionSuffixed.IsLabel);
            Assert.False(plainField.IsLabel);
        }
        finally
        {
            LabelConvention.Current = previous;
        }
    }

    [Fact]
    public void Labels_data_item_rule_is_enabled_by_default()
    {
        // The default convention enables the Labels-data-item rule (name "Labels"): the rule is
        // self-scoping - it can only ever match a document whose dataset actually carries a <Labels> data
        // item, and where that shape occurs its columns ARE captions (InventoryOrderDetails.docx is the
        // corpus proof; SchemaProviderTests/LayoutReaderTests pin its classification under this default).
        // Opting out is explicit: BCWL_LABELS_DATA_ITEM="-" on the host, labelsDataItemName: null in code.
        // See LabelConvention's own remarks for the full rationale.
        Assert.Equal("Labels", LabelConvention.Default.LabelsDataItemName);

        var column = new DatasetColumn { Name = "DataRetrieved", Path = "/Labels/DataRetrieved" };
        Assert.True(LabelConvention.Default.IsLabelPath(column.Path));

        // Self-scoping: the same unsuffixed name under any OTHER data item stays a plain field.
        Assert.False(LabelConvention.Default.IsLabelPath("/Header/DataRetrieved"));
    }
}
