using BcWordLayout.McpHost;
using BcWordLayout.McpHost.Tools;

namespace BcWordLayout.Tests;

/// <summary>
/// Cross-cutting coverage: drives every tool into its representative failure modes (missing file,
/// each distinct <c>invalid_argument</c> shape, <c>not_found</c> for a bad control id and a bad location
/// target, <c>invalid_layout</c> for a non-BC file) and asserts the STRUCTURAL guarantee added there -
/// <see cref="ToolError.Hint"/> is now non-nullable and <see cref="ToolResponse.Failure"/> requires a
/// <c>hint</c> argument (see their own remarks) - actually holds at runtime, and that the hint is genuinely
/// agent-actionable (not a generic restate): long enough to carry real guidance and, where applicable, naming
/// the offending argument or an inspection tool (<c>get_layout_info</c>/<c>list_dataset_fields</c>).
/// Complements, rather than duplicates, the per-tool error-path tests already in
/// <see cref="McpHostToolTests"/> (which mostly just assert the hint is non-empty).
/// </summary>
/// <remarks>Joins the preview-converter-seam collection because it calls
/// <c>LifecycleTools.PreviewLayout</c> (see <see cref="PreviewConverterSeamCollection"/> for the rule).</remarks>
[Collection("preview-converter-seam")]
public class ErrorHintTests
{
    // ---- shared assertion ----

    /// <summary>
    /// The shape every failure must have: <c>Ok=false</c>, no <c>Data</c>, a populated <c>Error</c> whose
    /// <c>Code</c> matches <paramref name="expectedCode"/>, a non-empty <c>Message</c>, and a <c>Hint</c>
    /// that is non-null (a compile-time guarantee as of - see <see cref="ToolError"/>'s remarks),
    /// non-whitespace, and long enough to be more than a bare restate. When <paramref name="hintMustMention"/>
    /// is supplied, the hint must also actually name the specific argument or inspection tool this failure
    /// mode calls for (case-insensitive) - proving the hint was TAILORED to this failure, not just present.
    /// </summary>
    private static void AssertActionableFailure(ToolResponse response, string expectedCode, string? hintMustMention = null)
    {
        Assert.False(response.Ok);
        Assert.Null(response.Data);
        Assert.NotNull(response.Error);

        var error = response.Error!;
        Assert.Equal(expectedCode, error.Code);
        Assert.False(string.IsNullOrWhiteSpace(error.Message));

        // Hint is `string` (non-nullable), so this is mostly already a compile-time
        // guarantee - a hintless ToolResponse.Failure call site cannot compile - but assert the runtime
        // value too rather than trusting the type system alone: whitespace-only text would defeat the whole
        // point of the guarantee just as much as null would, and is not something the type system alone
        // rules out.
        Assert.False(string.IsNullOrWhiteSpace(error.Hint));
        Assert.True(
            error.Hint.Trim().Length > 20,
            $"hint '{error.Hint}' for code '{expectedCode}' reads like a generic restate, not real guidance");

        if (hintMustMention is not null)
        {
            Assert.Contains(hintMustMention, error.Hint, StringComparison.OrdinalIgnoreCase);
        }
    }

    private static string CopyOfCorpus(string corpusFile)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-errorhint-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    private static string NewOutputPath() =>
        Path.Combine(Path.GetTempPath(), $"bcwl-errorhint-out-{Guid.NewGuid():N}.docx");

    // ---- file_not_found: the message names which parameter was checked, the hint says "absolute path" ----

    [Fact]
    public void GetLayoutInfo_missing_file_hint_is_actionable()
    {
        var response = ReadTools.GetLayoutInfo("Z:\\does-not-exist.docx");
        AssertActionableFailure(response, "file_not_found", "absolute path");
        Assert.Contains("layoutPath", response.Error!.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ListDatasetFields_missing_file_hint_is_actionable()
    {
        var response = ReadTools.ListDatasetFields("Z:\\does-not-exist.docx");
        AssertActionableFailure(response, "file_not_found", "absolute path");
        Assert.Contains("source", response.Error!.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ValidateLayout_missing_file_hint_is_actionable()
    {
        var response = ReadTools.ValidateLayout("Z:\\does-not-exist.docx");
        AssertActionableFailure(response, "file_not_found", "absolute path");
    }

    [Fact]
    public void PreviewLayout_missing_file_hint_is_actionable()
    {
        var response = LifecycleTools.PreviewLayout("Z:\\does-not-exist.docx");
        AssertActionableFailure(response, "file_not_found", "absolute path");
    }

    [Fact]
    public void CreateLayout_missing_schemaSource_hint_names_schemaSource()
    {
        var response = LifecycleTools.CreateLayout("Z:\\does-not-exist.docx", NewOutputPath());
        AssertActionableFailure(response, "file_not_found", "absolute path");
        Assert.Contains("schemaSource", response.Error!.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void CreateLayout_missing_templatePath_hint_names_templatePath()
    {
        var response = LifecycleTools.CreateLayout(
            Corpus.Path(Corpus.SalesInvoice), NewOutputPath(), "Z:\\does-not-exist.docx");
        AssertActionableFailure(response, "file_not_found", "absolute path");
        Assert.Contains("templatePath", response.Error!.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void InsertField_missing_file_hint_names_layoutPath()
    {
        var response = EditTools.InsertField("Z:\\does-not-exist.docx", "/Header/CustomerAddress1", "documentEnd");
        AssertActionableFailure(response, "file_not_found", "absolute path");
        Assert.Contains("layoutPath", response.Error!.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RemoveControl_missing_file_hint_is_actionable()
    {
        var response = EditTools.RemoveControl("Z:\\does-not-exist.docx", 1);
        AssertActionableFailure(response, "file_not_found", "absolute path");
    }

    [Fact]
    public void RefreshXmlPart_missing_layoutPath_hint_names_layoutPath()
    {
        var response = LifecycleTools.RefreshXmlPart("Z:\\does-not-exist.docx", Corpus.Path(Corpus.SalesInvoice));
        AssertActionableFailure(response, "file_not_found", "absolute path");
        Assert.Contains("layoutPath", response.Error!.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RefreshXmlPart_missing_newSchemaSource_hint_names_newSchemaSource()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = LifecycleTools.RefreshXmlPart(path, "Z:\\does-not-exist.xml");
            AssertActionableFailure(response, "file_not_found", "absolute path");
            Assert.Contains("newSchemaSource", response.Error!.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- invalid_layout: non-BC files - hint points at get_layout_info ----

    [Fact]
    public void CreateLayout_from_non_BC_xml_hint_points_at_get_layout_info()
    {
        var badXmlPath = Path.Combine(Path.GetTempPath(), $"bcwl-errorhint-badschema-{Guid.NewGuid():N}.xml");
        try
        {
            File.WriteAllText(badXmlPath, "<SomeOtherRoot xmlns=\"urn:not-bc\"><Foo/></SomeOtherRoot>");

            var response = LifecycleTools.CreateLayout(badXmlPath, NewOutputPath());

            AssertActionableFailure(response, "invalid_layout", "get_layout_info");
        }
        finally
        {
            File.Delete(badXmlPath);
        }
    }

    [Fact]
    public void RefreshXmlPart_invalid_newSchemaSource_hint_points_at_get_layout_info()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        var badXmlPath = Path.Combine(Path.GetTempPath(), $"bcwl-errorhint-refresh-badschema-{Guid.NewGuid():N}.xml");
        try
        {
            File.WriteAllText(badXmlPath, "<SomeOtherRoot xmlns=\"urn:not-bc\"><Foo/></SomeOtherRoot>");

            var response = LifecycleTools.RefreshXmlPart(path, badXmlPath);

            AssertActionableFailure(response, "invalid_layout", "get_layout_info");
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

    // ---- template_not_unbound: full BC layout as templatePath - hint routes to refresh_xml_part/remove_control ----

    [Fact]
    public void CreateLayout_with_a_full_BC_layout_templatePath_hint_routes_to_refresh_xml_part_and_remove_control()
    {
        var response = LifecycleTools.CreateLayout(
            Corpus.Path(Corpus.SalesInvoice), NewOutputPath(), Corpus.Path(Corpus.StandardStatement));

        AssertActionableFailure(response, "template_not_unbound", "refresh_xml_part");
        Assert.Contains("remove_control", response.Error!.Hint, StringComparison.OrdinalIgnoreCase);
    }

    // ---- invalid_argument: one representative test per distinct argument shape across the tool surface ----

    [Fact]
    public void ValidateLayout_bad_level_hint_names_valid_values()
    {
        var response = ReadTools.ValidateLayout(Corpus.Path(Corpus.SalesInvoice), "bogus");
        AssertActionableFailure(response, "invalid_argument", "quick");
    }

    [Fact]
    public void PreviewLayout_bad_converter_hint_names_valid_values()
    {
        var response = LifecycleTools.PreviewLayout(Corpus.Path(Corpus.SalesInvoice), converter: "bogus");
        AssertActionableFailure(response, "invalid_argument", "libreoffice");
    }

    [Fact]
    public void InsertField_bad_locationType_hint_names_valid_values()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = EditTools.InsertField(path, "/Header/CustomerAddress1", "not-a-real-location");
            AssertActionableFailure(response, "invalid_argument", "locationtype");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertField_afterControl_without_controlId_hint_names_controlId()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = EditTools.InsertField(path, "/Header/CustomerAddress1", "afterControl");
            AssertActionableFailure(response, "invalid_argument", "controlid");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertField_tableCell_missing_indices_hint_names_tableCell()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            // tableIndex supplied, row/col omitted - still incomplete for tableCell addressing.
            var response = EditTools.InsertField(path, "/Header/CustomerAddress1", "tableCell", tableIndex: 0);
            AssertActionableFailure(response, "invalid_argument", "tablecell");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertField_atText_without_searchText_hint_names_searchText()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = EditTools.InsertField(path, "/Header/CustomerAddress1", "atText");
            AssertActionableFailure(response, "invalid_argument", "searchtext");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertField_bad_dataset_path_hint_points_at_list_dataset_fields()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = EditTools.InsertField(path, "/Header/ThisFieldDoesNotExistAnywhere", "documentEnd");
            AssertActionableFailure(response, "invalid_argument", "list_dataset_fields");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertLabel_non_label_shaped_path_hint_points_at_list_dataset_fields()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            // Not label-shaped (no 'Lbl'/'_Lbl' suffix) - SdtFactory.BuildLabel rejects it up front.
            var response = EditTools.InsertLabel(path, "/Header/CustomerAddress1", "documentEnd");
            AssertActionableFailure(response, "invalid_argument", "list_dataset_fields");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertRepeaterTable_bad_dataItem_hint_points_at_list_dataset_fields()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = TableTools.InsertRepeaterTable(
                path, "/Header/ThisDataItemDoesNotExistAnywhere", "ItemNo_Line", "documentEnd");
            AssertActionableFailure(response, "invalid_argument", "list_dataset_fields");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertRepeaterTable_bad_column_hint_names_leaf_columns()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = TableTools.InsertRepeaterTable(
                path, "/Header/Line", "ThisColumnDoesNotExistAnywhere", "documentEnd");
            AssertActionableFailure(response, "invalid_argument", "leaf column");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertRepeaterTable_empty_columns_hint_names_leaf_columns()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = TableTools.InsertRepeaterTable(path, "/Header/Line", "   ", "documentEnd");
            AssertActionableFailure(response, "invalid_argument", "leaf column");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertRepeaterTable_columnWidths_mismatch_hint_names_columnWidths()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = TableTools.InsertRepeaterTable(
                path, "/Header/Line", "ItemNo_Line,Description_Line,TransHeaderAmount", "documentEnd",
                columnWidths: "1000,2000");
            AssertActionableFailure(response, "invalid_argument", "columnwidths");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RemoveControl_keepText_on_a_repeater_hint_names_keepText_false()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var info = Assert.IsType<LayoutInfoDto>(ReadTools.GetLayoutInfo(path).Data);
            var repeater = info.Controls.First(c => c.Kind == "Repeater" && c.SdtId.HasValue);

            var response = EditTools.RemoveControl(path, repeater.SdtId!.Value, keepText: true);

            AssertActionableFailure(response, "invalid_argument", "keeptext=false");
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- not_found: bad control id and bad location targets - hint points at get_layout_info ----

    [Fact]
    public void RemoveControl_of_a_nonexistent_id_hint_points_at_get_layout_info()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            // "not sequential or guessable" is distinctive to the Control-kind hint - every not_found hint
            // (including the generic fallback) mentions get_layout_info, so asserting THAT would not prove
            // the TargetKind-specific branch fired.
            var response = EditTools.RemoveControl(path, 999_999_999);
            AssertActionableFailure(response, "not_found", "not sequential or guessable");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertField_afterControl_with_nonexistent_id_hint_points_at_get_layout_info()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = EditTools.InsertField(
                path, "/Header/CustomerAddress1", "afterControl", controlId: 999_999_999);
            AssertActionableFailure(response, "not_found", "not sequential or guessable");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertField_tableCell_out_of_range_hint_points_at_get_layout_info()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = EditTools.InsertField(
                path, "/Header/CustomerAddress1", "tableCell", tableIndex: 999, row: 0, col: 0);
            AssertActionableFailure(response, "not_found", "all indices are 0-based");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertField_atText_not_present_hint_points_at_get_layout_info()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = EditTools.InsertField(
                path, "/Header/CustomerAddress1", "atText",
                searchText: "ThisTextIsDefinitelyNotInTheDocumentAnywhere");
            AssertActionableFailure(response, "not_found", "exact substring that is really present");
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- B11: NotFoundException.TargetKind drives the not_found hint - one more test per TargetKind not
    // already exercised above (Control/TableCoordinate/SearchText are pinned by the four tests above, each
    // asserting a fragment DISTINCTIVE to its kind-specific hint rather than one every hint shares; these
    // three cover NamedHeaderFooterPart, HeaderFooterParts, and AfterControlPosition), plus a test proving an
    // internal-bug-shaped InvalidOperationException is NOT reported as not_found (see NotFoundException's
    // own remarks: only a deliberately-constructed NotFoundException produces not_found; a raw
    // InvalidOperationException - the BCL's default failure type - now falls to internal_error instead). ----

    [Fact]
    public void InsertField_unknown_header_partName_hint_names_the_available_parts()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = EditTools.InsertField(
                path, "/Header/CustomerAddress1", "documentEnd",
                layoutPart: "header", partName: "no-such-header.xml");
            AssertActionableFailure(response, "not_found", "header1.xml");
            Assert.Contains("no-such-header.xml", response.Error!.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertField_header_targeted_on_a_layout_with_no_header_parts_hint_suggests_body()
    {
        // A synthetic layout has no header/footer parts at all - the NotFoundTarget.HeaderFooterParts
        // bucket, distinct from NamedHeaderFooterPart above (that layout DOES have header parts, just not
        // one matching the name). Deliberately an 'atText' location: 'documentEnd' now SCAFFOLDS the missing
        // part instead of failing (see LayoutEditorTests), while a location that could never
        // resolve inside a freshly created empty part still lands here.
        var path = SyntheticLayout.Create(SyntheticLayout.PlainParagraph("anchor"));
        try
        {
            var response = EditTools.InsertField(
                path, "/Header/CompanyName", "atText", searchText: "anchor", layoutPart: "header");
            AssertActionableFailure(response, "not_found", "body");
            Assert.Contains("no header parts", response.Error!.Message, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertField_afterControl_on_a_repeater_hint_suggests_tableCell_addressing()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var info = Assert.IsType<LayoutInfoDto>(ReadTools.GetLayoutInfo(path).Data);
            // Specifically a ROW-level repeater (its w:sdt sits directly inside a w:tbl, e.g. the line-items
            // table): AfterControl rejects that shape outright (a table cannot host a paragraph/block sdt as
            // a direct sibling) - unlike a block-level repeater (not inside a table), which AfterControl can
            // insert after just fine.
            var repeater = info.Controls.First(c => c.Kind == "Repeater" && c.SdtId.HasValue && c.Level == "row");

            var response = EditTools.InsertField(
                path, "/Header/CustomerAddress1", "afterControl", controlId: repeater.SdtId!.Value);

            AssertActionableFailure(response, "not_found", "tableCell");
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- internal_error: an internal-bug-shaped InvalidOperationException must NOT surface as not_found ----

    [Fact]
    public void CreateLayout_with_a_bare_drive_root_as_outputPath_returns_internal_error_not_not_found()
    {
        // LayoutBuilder.Create's "Could not determine the directory of ..." throw is a plain
        // InvalidOperationException (an internal invariant, not a lookup failure) - a bare drive root is the
        // one input that reaches it via the real public tool surface without any file ever being written
        // (the throw fires before Directory.CreateDirectory/any write). Confirms Guard's generic
        // catch (Exception) - not a not_found branch - is what handles it post-B11.
        var response = LifecycleTools.CreateLayout(Corpus.Path(Corpus.SalesInvoice), "C:\\");

        Assert.False(response.Ok);
        Assert.Equal("internal_error", response.Error!.Code);
        Assert.Contains("Could not determine the directory", response.Error!.Message, StringComparison.Ordinal);
    }
}
