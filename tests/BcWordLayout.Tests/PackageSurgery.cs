using System.IO.Compression;

namespace BcWordLayout.Tests;

/// <summary>
/// Byte-level edits to a COPY of a corpus layout's zip, for the package shapes no tool produces and the
/// OpenXml SDK would not write, such as a custom XML part whose bytes contradict its own encoding declaration
/// or a truncated part. Every helper works on a temp copy - the corpus itself is never opened for writing.
/// </summary>
internal static class PackageSurgery
{
    /// <summary>Copies <paramref name="corpusFile"/> to a fresh temp path and returns it.</summary>
    public static string CopyOfCorpus(string corpusFile, string tag = "surgery")
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-{tag}-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    /// <summary>The raw bytes of zip entry <paramref name="entryName"/> (e.g. <c>customXml/item1.xml</c>).</summary>
    public static byte[] ReadEntry(string docxPath, string entryName)
    {
        using var zip = ZipFile.OpenRead(docxPath);
        var entry = zip.GetEntry(entryName) ?? throw new InvalidOperationException($"No entry '{entryName}' in '{docxPath}'.");
        using var stream = entry.Open();
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }

    /// <summary>Replaces zip entry <paramref name="entryName"/>'s content with <paramref name="bytes"/>, verbatim.</summary>
    public static void ReplaceEntry(string docxPath, string entryName, byte[] bytes)
    {
        using var zip = ZipFile.Open(docxPath, ZipArchiveMode.Update);
        zip.GetEntry(entryName)?.Delete();
        var entry = zip.CreateEntry(entryName);
        using var stream = entry.Open();
        stream.Write(bytes, 0, bytes.Length);
    }
}
