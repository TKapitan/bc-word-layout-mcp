namespace BcWordLayout.Render;

/// <summary>
/// Shared output sanity-check used by every <see cref="IPdfConverter"/> once it believes a conversion
/// succeeded — a converter reporting <c>Ok = true</c> over a file that is not actually a valid PDF would be
/// a worse failure mode than reporting <c>Ok = false</c>, since the caller would only find out downstream.
/// </summary>
internal static class PdfFileValidation
{
    private static readonly byte[] Magic = System.Text.Encoding.ASCII.GetBytes("%PDF");

    /// <summary>True if <paramref name="path"/> exists, is non-empty, and starts with the <c>%PDF</c> magic number.</summary>
    internal static bool LooksLikePdf(string path)
    {
        try
        {
            if (!File.Exists(path))
            {
                return false;
            }

            using var stream = File.OpenRead(path);
            if (stream.Length < Magic.Length)
            {
                return false;
            }

            var buffer = new byte[Magic.Length];
            var read = stream.Read(buffer, 0, buffer.Length);
            return read == Magic.Length && buffer.AsSpan().SequenceEqual(Magic);
        }
        catch (IOException)
        {
            return false;
        }
        catch (UnauthorizedAccessException)
        {
            return false;
        }
    }
}
