namespace BcWordLayout.Tests;

/// <summary>
/// Serializes every test class that sets <c>ToolGuards.BcVersionEnvironmentValue</c>, the stand-in for the
/// server's <c>BCWL_BC_VERSION</c> (GitHub issue #152), or asserts the target a tool call resolves: every tool call
/// reads the seam, so a class asserting an unknown target must never run while another has set it.
/// </summary>
[CollectionDefinition("bc-version-environment-seam")]
public sealed class BcVersionEnvironmentSeamCollection
{
    // Marker class: holds the [CollectionDefinition] attribute only; xUnit matches members by the
    // collection NAME string on each [Collection] attribute, never by this type.
}
