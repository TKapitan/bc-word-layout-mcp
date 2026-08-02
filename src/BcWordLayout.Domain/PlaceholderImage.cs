namespace BcWordLayout.Domain;

/// <summary>
/// A tiny, genuinely valid placeholder PNG for a picture content control's image — real BC layouts ship a
/// 10-byte stub <c>image1.bin</c> instead, which is not a decodable image at all. A picture control's
/// on-page size comes from its drawing's extent, not the source image's pixel dimensions, so a small solid
/// bitmap scales to fill whatever frame the layout defines. Two callers share it, which is why it lives in
/// Domain rather than beside either: <c>BcWordLayout.Merge</c>'s merge engine repoints an EXISTING picture
/// control's blip at it so a preview has something rasterizable, and <see cref="LayoutEditor.InsertPicture"/>
/// embeds it as the initial content of a NEWLY authored picture placeholder (a blip must point at a real
/// image part for Word to open the document, and BC replaces the image at render time anyway).
/// </summary>
public static class PlaceholderImage
{
    /// <summary>
    /// Raw bytes of a 32x32, solid light-gray (<c>#CCCCCC</c>) PNG. Starts with the 8-byte PNG signature
    /// (<c>89 50 4E 47 0D 0A 1A 0A</c>) and contains exactly one <c>IHDR</c>, one <c>IDAT</c>, and one
    /// <c>IEND</c> chunk, each with a verified CRC — confirmed by round-tripping this exact constant
    /// through a PNG decoder when it was generated (98 bytes total).
    /// </summary>
    public static readonly byte[] PngBytes = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAIAAAD8GO2jAAAAKUlEQVR42u3NQQEAAAQEMPSPdqGU4LcVWCepT1PPBAKBQCAQ"
        + "CAQCwZUFiZoCpGJJ0gQAAAAASUVORK5CYII=");
}
