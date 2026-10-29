using BcWordLayout.Domain;
using BcWordLayout.McpHost;
using BcWordLayout.McpHost.Tools;

namespace BcWordLayout.Tests;

/// <summary>
/// GitHub issue #93: every tool that writes a binding took its storeItemID from the BC dataset part's
/// itemProps, so on a layout whose part has none — 4 of the 13 corpus layouts, among them the stock
/// SalespersonCommission and StandardStatement — every bound insert failed with an invalid_argument whose hint
/// named internal C# methods. Those layouts' existing bindings agree on one storeItemID (45 of 45 and 82 of
/// 82), and a new binding now reuses it; with no bindings, or disagreeing ones, the call is refused with a hint
/// saying what the caller can do. No itemProps is ever minted.
/// </summary>
public class StoreItemIdFallbackTests
{
    private const string CommissionStoreItemId = "{F622BF70-BE75-4E42-912A-021D50085D69}";
    private const string StatementStoreItemId = "{55AAF175-0B56-41B3-BB08-59F8774CF590}";

    /// <summary>The storeItemID the control <paramref name="controlId"/> was written with.</summary>
    private static string? StoreItemIdOf(string path, int controlId) =>
        Assert.IsType<LayoutInfoDto>(ReadTools.GetLayoutInfo(path).Data).Controls.Single(c => c.SdtId == controlId).StoreItemId;

    private static void AssertNoItemPropsMinted(string path)
    {
        using var doc = LayoutPackage.Open(path, isEditable: false);
        Assert.All(doc.MainDocumentPart!.CustomXmlParts, p => Assert.Null(p.CustomXmlPropertiesPart));
    }

    [Fact]
    public void SalespersonCommission_insert_field_and_insert_label_reuse_the_bindings_storeItemId()
    {
        var path = PackageSurgery.CopyOfCorpus(Corpus.SalespersonCommission, "storeitem");
        try
        {
            var errorsBefore = LayoutValidator.Quick(path).ErrorCount;
            var label = SchemaProvider.FromLayout(path).AllColumns().First(c => c.IsLabel).Path;

            var field = EditTools.InsertField(path, "/Salesperson_Purchaser/Salesperson_Purchaser_Code", "documentEnd");
            var labelResponse = EditTools.InsertLabel(path, label, "documentEnd");

            Assert.True(field.Ok, field.Error?.Message);
            Assert.True(labelResponse.Ok, labelResponse.Error?.Message);
            Assert.Equal(CommissionStoreItemId, StoreItemIdOf(path, Assert.IsType<EditResultDto>(field.Data).ControlId));
            Assert.Equal(CommissionStoreItemId, StoreItemIdOf(path, Assert.IsType<EditResultDto>(labelResponse.Data).ControlId));
            // The table-structure tools build their bound cells through the same lookup.
            var info = Assert.IsType<LayoutInfoDto>(ReadTools.GetLayoutInfo(path).Data);
            var boundBefore = info.Controls.Count(c => c.StoreItemId == CommissionStoreItemId);
            var table = info.Tables[0];
            var row = TableTools.InsertTableRow(
                path, table.TableIndex, [new() { Fields = ["/Salesperson_Purchaser/Salesperson_Purchaser_Code"], Span = table.ColumnCount }]);
            Assert.True(row.Ok, row.Error?.Message);
            Assert.Equal(
                boundBefore + 1,
                Assert.IsType<LayoutInfoDto>(ReadTools.GetLayoutInfo(path).Data).Controls.Count(c => c.StoreItemId == CommissionStoreItemId));

            Assert.Equal(errorsBefore, LayoutValidator.Quick(path).ErrorCount);
            AssertNoItemPropsMinted(path);

            // get_layout_info still reports what the PART declares: no itemProps, so no storeItemID.
            Assert.Null(Assert.IsType<LayoutInfoDto>(ReadTools.GetLayoutInfo(path).Data).Report.StoreItemId);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void StandardStatement_insert_field_and_insert_picture_reuse_the_bindings_storeItemId()
    {
        var path = PackageSurgery.CopyOfCorpus(Corpus.StandardStatement, "storeitem");
        try
        {
            var errorsBefore = LayoutValidator.Quick(path).ErrorCount;
            var picture = SchemaProvider.FromLayout(path).AllColumns().First(c => c.Name == "CompanyPicture").Path;

            var field = EditTools.InsertField(path, "/Customer/No_Cust", "documentEnd");
            var pictureResponse = EditTools.InsertPicture(path, picture, "documentEnd");

            Assert.True(field.Ok, field.Error?.Message);
            Assert.True(pictureResponse.Ok, pictureResponse.Error?.Message);
            Assert.Equal(StatementStoreItemId, StoreItemIdOf(path, Assert.IsType<EditResultDto>(field.Data).ControlId));
            Assert.Equal(StatementStoreItemId, StoreItemIdOf(path, Assert.IsType<EditResultDto>(pictureResponse.Data).ControlId));
            Assert.Equal(errorsBefore, LayoutValidator.Quick(path).ErrorCount);
            AssertNoItemPropsMinted(path);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Bindings_that_differ_only_in_braces_and_case_agree_and_the_first_literal_is_reused()
    {
        const string xpath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:CompanyName[1]";
        var path = SyntheticLayout.CreateWithoutStoreItemId(
            SyntheticLayout.BoundField(xpath, "{ABCDEF01-0000-0000-0000-000000000001}")
            + SyntheticLayout.BoundField(xpath, "abcdef01-0000-0000-0000-000000000001"));
        try
        {
            var response = EditTools.InsertField(path, "/Header/CompanyName", "documentEnd");

            Assert.True(response.Ok, response.Error?.Message);
            Assert.Equal(
                "{ABCDEF01-0000-0000-0000-000000000001}",
                StoreItemIdOf(path, Assert.IsType<EditResultDto>(response.Data).ControlId));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void PaymentPractice_bindings_disagree_so_a_bound_insert_is_refused_with_an_actionable_hint()
    {
        // 20 bindings name one storeItemID and 5 another, and the part has no itemProps: reusing either
        // would be a guess.
        var path = PackageSurgery.CopyOfCorpus(Corpus.PaymentPracticeByPeriod, "storeitem");
        try
        {
            var before = File.ReadAllBytes(path);
            var column = SchemaProvider.FromLayout(path).AllColumns().First(c => !c.IsLabel).Path;

            var response = EditTools.InsertField(path, column, "documentEnd");

            Assert.False(response.Ok);
            Assert.Equal("invalid_layout", response.Error!.Code);
            Assert.Contains("carries no itemProps, so it declares no storeItemID", response.Error.Message);
            Assert.Contains("the layout's 25 existing dataset bindings do not agree on one", response.Error.Message);
            Assert.Contains("(20x)", response.Error.Message);
            Assert.Contains("(5x)", response.Error.Message);
            Assert.Contains("upload the layout to Business Central and export it back", response.Error.Hint);
            Assert.Contains("whose existing bindings all name the same storeItemID", response.Error.Hint);
            foreach (var internalName in new[] { "SchemaProvider", "FromLayout", "FromSchemaXml", "SdtFactory" })
            {
                Assert.DoesNotContain(internalName, response.Error.Message, StringComparison.Ordinal);
                Assert.DoesNotContain(internalName, response.Error.Hint, StringComparison.Ordinal);
            }

            Assert.Equal(before, File.ReadAllBytes(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_layout_with_no_itemProps_and_no_bindings_is_refused_by_every_binding_tool()
    {
        var path = SyntheticLayout.CreateWithoutStoreItemId(string.Empty, SyntheticLayout.DatasetWithLineXml);
        try
        {
            var responses = new[]
            {
                EditTools.InsertField(path, "/Header/CompanyName", "documentEnd"),
                EditTools.InsertPicture(path, "/Header/CompanyName", "documentEnd"),
                TableTools.InsertRepeaterTable(path, "/Header/Line", TableArgs.FieldsOf("/Header/Line", "ItemNo"), "documentEnd"),
            };

            Assert.All(responses, r =>
            {
                Assert.Equal("invalid_layout", r.Error!.Code);
                Assert.Contains("this layout has no existing dataset binding whose storeItemID could be reused", r.Error.Message);
                Assert.Contains("A new itemProps is deliberately not invented here", r.Error.Hint);
            });

            // Unbound content is not a binding and needs no storeItemID.
            Assert.True(EditTools.InsertText(path, "Statement", "documentEnd").Ok);
            AssertNoItemPropsMinted(path);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
