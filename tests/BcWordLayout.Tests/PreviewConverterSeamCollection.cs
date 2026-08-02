namespace BcWordLayout.Tests;

/// <summary>
/// Serializes every test class that can REACH the process-wide <c>LifecycleTools.SelectConverter</c> seam
///: xUnit runs classes in the same named collection sequentially, so a test swapping the seam
/// for a <c>FakePdfConverter</c> can never overlap a test in another class whose <c>preview_layout</c> call
/// would read it mid-swap and resolve the fake instead of the real factory (an intermittent, machine-order-
/// dependent failure). Membership rule - stated on <c>SelectConverter</c>'s own doc comment too: any class
/// with a test that calls <c>LifecycleTools.PreviewLayout</c> AT ALL joins this collection (strictly only
/// calls that get past the file-exists and converter-kind argument gates read the seam, but "calls the tool
/// at all" is the rule a contributor can apply without tracing the tool body); classes that bypass the tool
/// entirely (e.g. <c>FidelityHarnessTests</c>, which drives <c>MergeEngine</c> + <c>PdfConverterFactory</c>
/// directly) never read the seam and need not join.
/// </summary>
[CollectionDefinition("preview-converter-seam")]
public sealed class PreviewConverterSeamCollection
{
    // Marker class: holds the [CollectionDefinition] attribute only; xUnit matches members by the
    // collection NAME string on each [Collection] attribute, never by this type.
}
