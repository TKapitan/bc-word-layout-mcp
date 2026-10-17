namespace BcWordLayout.Tests;

/// <summary>
/// Locates files in the repository checkout for the source-scan and documentation guard tests (the same walk
/// up from the test assembly that <c>PackageOpenGatewayTests</c> uses). Tests that need a file at a fixed
/// spot next to the test assembly copy it there instead (see the <c>docsync</c> and <c>packaging</c> links in
/// the test project file).
/// </summary>
internal static class RepoFiles
{
    private static readonly Lazy<string> RootDirectory = new(FindRoot);

    /// <summary>The checkout's root: the folder holding <c>bc-word-layout-mcp.sln</c>.</summary>
    public static string Root => RootDirectory.Value;

    /// <summary>The full path of <paramref name="relativePath"/> (forward slashes allowed) under <see cref="Root"/>.</summary>
    public static string PathOf(string relativePath) =>
        Path.Combine(Root, relativePath.Replace('/', Path.DirectorySeparatorChar));

    /// <summary>The text of <paramref name="relativePath"/> under <see cref="Root"/>.</summary>
    public static string Read(string relativePath) => File.ReadAllText(PathOf(relativePath));

    /// <summary>Every <c>*.cs</c> file under <paramref name="relativeFolder"/>, build output excluded.</summary>
    public static IEnumerable<string> SourceFiles(string relativeFolder) =>
        Directory.EnumerateFiles(PathOf(relativeFolder), "*.cs", SearchOption.AllDirectories).Where(f => !IsBuildOutput(f));

    /// <summary><paramref name="fullPath"/> relative to <see cref="Root"/>, with forward slashes, for messages.</summary>
    public static string Relative(string fullPath) => Path.GetRelativePath(Root, fullPath).Replace('\\', '/');

    private static bool IsBuildOutput(string path) =>
        path.Contains($"{Path.DirectorySeparatorChar}bin{Path.DirectorySeparatorChar}", StringComparison.Ordinal)
        || path.Contains($"{Path.DirectorySeparatorChar}obj{Path.DirectorySeparatorChar}", StringComparison.Ordinal);

    private static string FindRoot()
    {
        for (var dir = new DirectoryInfo(AppContext.BaseDirectory); dir is not null; dir = dir.Parent)
        {
            if (File.Exists(Path.Combine(dir.FullName, "bc-word-layout-mcp.sln")))
            {
                return dir.FullName;
            }
        }

        throw new DirectoryNotFoundException("Could not find the repository root above the test assembly.");
    }
}
