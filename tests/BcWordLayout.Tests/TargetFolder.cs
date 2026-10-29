namespace BcWordLayout.Tests;

/// <summary>
/// A temp folder standing in for a repository, for tests of the target BC version (GitHub issue #152): it carries
/// a <c>.git</c> folder, so the walk up from a layout inside it never reaches an <c>app.json</c> outside it.
/// Deleted on dispose.
/// </summary>
internal sealed class TargetFolder : IDisposable
{
    public TargetFolder()
    {
        Root = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"bcwl-target-{Guid.NewGuid():N}");
        Directory.CreateDirectory(System.IO.Path.Combine(Root, ".git"));
    }

    public string Root { get; }

    /// <summary>
    /// The absolute path of <paramref name="relativePath"/> (written with <c>/</c>) under <see cref="Root"/>, in the
    /// platform's own form; its folder is created.
    /// </summary>
    public string Path(string relativePath)
    {
        var full = System.IO.Path.GetFullPath(System.IO.Path.Combine(Root, relativePath));
        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(full)!);
        return full;
    }

    /// <summary>Writes an AL <c>app.json</c> into <paramref name="relativeFolder"/>, stating <paramref name="application"/>.</summary>
    public string AppJson(string relativeFolder, string application) =>
        Write(System.IO.Path.Combine(relativeFolder, "app.json"), $$"""{ "name": "Layouts", "application": "{{application}}" }""");

    public string Write(string relativePath, string content)
    {
        var path = Path(relativePath);
        File.WriteAllText(path, content);
        return path;
    }

    /// <summary>Copies <paramref name="sourcePath"/> to <paramref name="relativePath"/> under <see cref="Root"/>.</summary>
    public string CopyIn(string sourcePath, string relativePath)
    {
        var path = Path(relativePath);
        File.Copy(sourcePath, path, overwrite: true);
        return path;
    }

    public void Dispose()
    {
        try
        {
            Directory.Delete(Root, recursive: true);
        }
        catch (IOException)
        {
            // Best effort: a temp folder left behind must not fail the test that used it.
        }
    }
}
