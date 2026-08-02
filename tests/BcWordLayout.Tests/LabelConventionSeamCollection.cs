namespace BcWordLayout.Tests;

/// <summary>
/// Serializes every test class that can observe <c>LabelConvention.Current</c> being
/// temporarily swapped for a custom convention: xUnit runs classes in the same named collection
/// sequentially, so a test installing a widened convention (e.g. suffixes <c>["Lbl", "Caption"]</c>) can
/// never overlap a test in another class whose label/field classification would read it mid-swap and get a
/// different answer than the default convention gives (an intermittent, machine-order-dependent failure).
/// </summary>
/// <remarks>
/// Membership rule: any class with a test whose assertions depend on label/field classification (directly
/// via <c>LabelConvention</c>/<c>DatasetColumn.IsLabel</c>/<c>ControlKind.Label</c>, or indirectly via
/// merge/sample-data generation's label-vs-field value strategy) for a schema/layout containing a
/// non-"Lbl"-suffixed label-shaped column joins this collection. In the current corpus that is exactly one
/// file — <c>InventoryOrderDetails.docx</c>, whose label-like columns are suffixed <c>Caption</c>/<c>Label</c>
/// (see <c>LabelConvention</c>'s own remarks) — so a class needs to join only if it asserts something
/// classification-sensitive about THAT file specifically:
/// <list type="bullet">
/// <item><see cref="SchemaProviderTests"/> — pins <c>InventoryOrderDetails</c>'s <c>IsLabel</c>-column
/// classification (exactly the <c>&lt;Labels&gt;</c> item's direct columns, via the default convention's
/// labels-data-item rule).</item>
/// <item><see cref="LayoutReaderTests"/> — pins <c>InventoryOrderDetails</c>'s
/// <see cref="BcWordLayout.Domain.Models.ControlKind.Label"/>-control classification (exactly the controls
/// bound into the <c>&lt;Labels&gt;</c> item).</item>
/// <item><see cref="MergeSnapshotTests"/> — byte-level merge-output snapshot for <c>InventoryOrderDetails</c>, whose text
/// content differs for label vs. field columns (see <c>SampleDataGenerator.GenerateLeafValue</c>).</item>
/// </list>
/// Every other corpus file's label-like columns already end in <c>Lbl</c>/<c>_Lbl</c> and so remain
/// classified as labels under any convention this repo's tests install (which always keeps <c>"Lbl"</c> in
/// the suffix list), so classes that only ever touch those files need not join. <c>InventoryOrderDetails</c>
/// is different because its classification hinges on the labels-data-item rule (on by default, but swapped
/// conventions may omit it — e.g. <c>new LabelConvention(["Lbl", "Caption"])</c> disables it). Tests that
/// swap <c>LabelConvention.Current</c> itself MUST also live in this collection and restore it in a
/// <c>finally</c> block regardless of outcome (see <see cref="LabelConventionConfigTests"/>).
/// </remarks>
[CollectionDefinition("label-convention-seam")]
public sealed class LabelConventionSeamCollection
{
    // Marker class: holds the [CollectionDefinition] attribute only; xUnit matches members by the
    // collection NAME string on each [Collection] attribute, never by this type.
}
