using System.Xml;
using System.Xml.Linq;

namespace BcWordLayout.Domain;

/// <summary>
/// Central resource ceilings guarding this repo's own hand-rolled XML loaders and recursive tree walkers
/// against a maliciously crafted or merely corrupted <c>.docx</c>/schema <c>.xml</c>. Two distinct crash
/// classes are guarded, both of which kill the host process rather than failing a call:
/// <list type="bullet">
/// <item><b>Unbounded load → OOM.</b> A part that is a few KB compressed in the zip can expand to
/// gigabytes of repetitive XML; loading it whole into an <c>XDocument</c>/OpenXML DOM with no size ceiling
/// exhausts memory. Guarded by the byte/part-count caps below.</item>
/// <item><b>Unbounded recursion → <see cref="StackOverflowException"/>.</b> The hand-rolled recursive
/// walkers over a schema/document tree have no natural bottom, and a crafted file nested tens of thousands
/// of levels deep overflows the stack. A <see cref="StackOverflowException"/> CANNOT be caught, so the
/// whole server dies. Guarded by the depth caps below, which fail as an ordinary catchable
/// exception instead.</item>
/// </list>
/// <para>
/// Every cap here is deliberately generous relative to a REAL Business Central layout/schema: a real BC
/// dataset custom XML part is documented at a few KB, and its data-item nesting at 3-4 levels (see the
/// <c>bc-word-layout-ooxml-facts</c> project memory) — so no legitimate input can ever trip one of these; a
/// tripped cap is itself diagnostic that the input is malformed or adversarial. Every violation is reported
/// as a dedicated <see cref="ResourceLimitExceededException"/> (a typed marker, not an
/// <see cref="InvalidDataException"/> — that BCL type is <c>sealed</c>) that <c>ToolGuards.Guard</c> catches
/// explicitly and maps to the SAME <c>invalid_layout</c> envelope every other malformed-layout finding gets
/// (so a crafted file surfaces as a normal user-input error, never an <c>internal_error</c>/process crash),
/// but with its OWN tailored hint — "this file is too big/too deeply nested" rather than the generic
/// "missing dataset part/wrong namespace" one (see
/// <see cref="ResourceLimitExceededException"/>'s own remarks).
/// </para>
/// <para>
/// Explicitly OUT OF SCOPE: the OpenXml SDK's own <c>MainDocumentPart.Document</c> (and header/footer) load
/// — that parse is SDK-internal, not this repo's code, and is unaffected by anything here.
/// </para>
/// </summary>
internal static class ResourceLimits
{
    /// <summary>
    /// Byte ceiling for a single custom XML part / standalone schema XML file this repo reads directly —
    /// see <see cref="SchemaProvider.FindBcParts"/> (every custom XML part inside a <c>.docx</c>),
    /// <see cref="SchemaProvider.FromSchemaXml"/> (a standalone exported schema <c>.xml</c>), the data-
    /// overrides loader in <c>BcWordLayout.Merge.SampleDataGenerator</c>, and the raw byte-for-byte part
    /// copies in <see cref="LayoutBuilder"/>/<see cref="LayoutRefresher"/>.
    /// <para>
    /// The rationale differs by caller. For the BC dataset SCHEMA part / a standalone schema <c>.xml</c>, 16
    /// MB is thousands of times a real part's actual size (documented at a few KB) — generous headroom for
    /// even an unusually large real export while still bounding a compressed-in-the-zip "bomb" part's
    /// decompressed size to a small, always-survivable allocation. For
    /// <c>BcWordLayout.Merge.SampleDataOptions.DataOverridesPath</c>, though, that "real ones are KBs"
    /// argument does NOT apply — an override is a full exported BC REPORT DATASET (potentially many repeater
    /// rows), not a schema, so a legitimately large real export could plausibly approach or exceed this cap.
    /// There the same 16 MB figure is instead a deliberate PREVIEW-SCALE budget: this tool exists to preview/
    /// validate a layout with representative sample data, not to merge arbitrarily large production exports,
    /// so capping how much of an override this repo will load is an acceptable, documented limitation rather
    /// than a bug — a caller with a genuinely larger real dataset should trim it before using
    /// <c>DataOverridesPath</c>. Kept as ONE shared constant (rather than a second, larger one for overrides)
    /// for simplicity; revisit only if a real override legitimately needs to exceed it.
    /// </para>
    /// </summary>
    internal const long MaxCustomXmlPartBytes = 16L * 1024 * 1024;

    /// <summary>
    /// Hard ceiling on the NUMBER of custom XML parts <see cref="SchemaProvider.FindBcPart"/>/
    /// <see cref="SchemaProvider.FindBcParts"/> will even attempt to enumerate. The per-part byte cap above
    /// bounds any ONE part's cost, but a package could still carry an enormous NUMBER of small,
    /// cheap-to-compress parts, making the per-part iteration/parse LOOP itself the unbounded cost - the byte
    /// cap alone does not bound it. A real BC layout carries a
    /// handful of custom XML parts (the BC dataset part, occasionally an unrelated Office bibliography part);
    /// 1024 is generous headroom while still bounding the loop to a small, fast, always-survivable amount of
    /// work.
    /// </summary>
    internal const int MaxCustomXmlParts = 1024;

    /// <summary>
    /// <see cref="XmlReaderSettings.MaxCharactersInDocument"/> paired with <see cref="MaxCustomXmlPartBytes"/>
    /// as a second, XML-parser-level ceiling (characters, not raw bytes) on every capped load — belt-and-
    /// braces alongside the length-limiting stream every capped load already goes through (see
    /// <see cref="LoadXDocumentCapped"/>). Text is at least 1 byte/char, so this can never itself be the
    /// binding constraint once the byte cap above is in force; it costs nothing to also set.
    /// </summary>
    internal const long MaxXmlCharacters = MaxCustomXmlPartBytes;

    /// <summary>
    /// <see cref="XmlReaderSettings.MaxCharactersFromEntities"/> for every capped load. This repo's XML is
    /// never legitimately DTD/entity-bearing (BC custom XML/schema parts are plain exported data), and
    /// <see cref="XmlReaderSettings.DtdProcessing"/> is already set to <see cref="DtdProcessing.Prohibit"/>
    /// on the same settings (which already throws on any <c>DOCTYPE</c> before an entity could even be
    /// declared) — this is defensive belt-and-braces only, not the primary XXE/entity-expansion defense.
    /// </summary>
    internal const long MaxCharactersFromEntities = 1024;

    /// <summary>
    /// Max recursion depth for the hand-rolled walkers over the parsed SCHEMA/dataset tree —
    /// <see cref="SchemaProvider.BuildNode"/>, the one place a <see cref="Models.DataItem"/> tree is ever
    /// constructed. <c>BcWordLayout.Merge.SampleDataGenerator.BuildInstance</c> and
    /// <c>BcWordLayout.McpHost.Tools.ToolGuards.ToDataItemDto</c> both recurse over that SAME tree
    /// afterward, so capping construction here structurally bounds their recursion too, transitively, with
    /// no separate counter needed at either (see each method's own remarks). Real BC datasets are documented
    /// at 3-4 levels; 64 is generous headroom while keeping every one of these small-per-frame methods well
    /// inside the default 1 MB thread stack even at the cap.
    /// </summary>
    internal const int MaxSchemaDepth = 64;

    /// <summary>
    /// Max recursion depth for the hand-rolled walkers over the OOXML DOCUMENT element tree —
    /// <c>BcWordLayout.Merge.MergeEngine.WalkElement</c> and <see cref="LayoutReader.Walk"/> — each of which
    /// recurses once per child element/content-control nesting level of document.xml (or a header/footer
    /// part), a shape this repo does not otherwise bound (unlike the schema tree above, real document
    /// nesting has no schema-derived ceiling). A real layout's paragraph/table/content-control nesting is a
    /// handful of levels; 128 is generous headroom.
    /// </summary>
    internal const int MaxElementNestingDepth = 128;

    /// <summary>
    /// Max recursion depth for the hand-rolled walk over the OPC PACKAGE's part-relationship graph —
    /// <see cref="LayoutValidator"/>'s external-relationship enumeration and
    /// <c>BcWordLayout.Merge.ExternalRelationshipStripper.StripPart</c>. Both already guard against a CYCLE
    /// (a visited-set skips any part already seen), but a crafted package could still chain many thousands
    /// of otherwise-acyclic parts to make the recursive walk itself arbitrarily deep. A real <c>.docx</c>'s
    /// part graph (document → header/footer/settings/glossary, rarely nested further) is a handful of
    /// levels; 128 is generous headroom.
    /// </summary>
    internal const int MaxPartGraphDepth = 128;

    /// <summary>
    /// Builds the <see cref="ResourceLimitExceededException"/> every depth-guarded walker throws once
    /// <paramref name="limit"/> is exceeded — one shared wording so every guard reads the same way regardless
    /// of which walker tripped it.
    /// </summary>
    internal static ResourceLimitExceededException DepthExceeded(string walkerName, int limit) =>
        new($"{walkerName} nesting exceeds the supported depth of {limit}; this looks like a malformed or "
            + "maliciously crafted file rather than a real Business Central layout/schema.");

    /// <summary>
    /// Builds the <see cref="ResourceLimitExceededException"/> thrown when a package carries more than
    /// <paramref name="limit"/> (<see cref="MaxCustomXmlParts"/>) custom XML parts — before any of them are
    /// even opened/parsed (see <see cref="SchemaProvider"/>'s own <c>EnsurePartCountWithinLimit</c>).
    /// </summary>
    internal static ResourceLimitExceededException PartCountExceeded(int actualCount, int limit) =>
        new($"Package has {actualCount} custom XML parts, exceeding the supported limit of {limit}; this "
            + "looks like a malformed or maliciously crafted file rather than a real Business Central layout.");

    /// <summary>
    /// Loads <paramref name="stream"/> as an <see cref="XDocument"/> through a length-limiting wrapper
    /// (<see cref="MaxCustomXmlPartBytes"/>, the PRIMARY defense — counts bytes actually read rather than
    /// trusting a possibly-absent/misleading <see cref="Stream.Length"/>) plus a capped, DTD-prohibited
    /// <see cref="XmlReader"/> (<see cref="MaxXmlCharacters"/>/<see cref="MaxCharactersFromEntities"/> —
    /// belt-and-braces). <paramref name="description"/> names what is being loaded (e.g. the part's file
    /// name) for the <see cref="ResourceLimitExceededException"/> message thrown when either cap is exceeded.
    /// </summary>
    internal static XDocument LoadXDocumentCapped(Stream stream, string description)
    {
        using var limited = new LengthLimitedStream(stream, MaxCustomXmlPartBytes, description);
        using var reader = XmlReader.Create(limited, CreateSafeSettings());
        return XDocument.Load(reader);
    }

    /// <summary>
    /// Reads the whole of <paramref name="source"/> into a byte array through the SAME length-limiting
    /// wrapper <see cref="LoadXDocumentCapped"/> uses (<see cref="MaxCustomXmlPartBytes"/>) — for callers
    /// (<see cref="LayoutBuilder"/>/<see cref="LayoutRefresher"/>) that need a custom XML part's RAW bytes
    /// verbatim rather than a parsed tree.
    /// </summary>
    internal static byte[] ReadAllBytesCapped(Stream source, string description)
    {
        using var limited = new LengthLimitedStream(source, MaxCustomXmlPartBytes, description);
        using var buffer = new MemoryStream();
        limited.CopyTo(buffer);
        return buffer.ToArray();
    }

    private static XmlReaderSettings CreateSafeSettings() => new()
    {
        DtdProcessing = DtdProcessing.Prohibit,
        XmlResolver = null,
        MaxCharactersInDocument = MaxXmlCharacters,
        MaxCharactersFromEntities = MaxCharactersFromEntities,
    };

    /// <summary>
    /// A read-only wrapper throwing <see cref="ResourceLimitExceededException"/> the moment more than
    /// <paramref name="maxBytes"/> bytes have been read from the wrapped stream — the actual enforcement
    /// mechanism behind <see cref="LoadXDocumentCapped"/>/<see cref="ReadAllBytesCapped"/>. Deliberately
    /// counts bytes ACTUALLY READ rather than trusting <see cref="Stream.Length"/> upfront: not every OPC
    /// part stream reliably reports one, and this way the cap holds regardless of what the stream claims.
    /// Never disposes the wrapped stream — ownership stays with the caller, exactly like every existing
    /// <c>using var stream = part.GetStream(...)</c> call site this wraps.
    /// </summary>
    private sealed class LengthLimitedStream : Stream
    {
        private readonly Stream _inner;
        private readonly long _maxBytes;
        private readonly string _description;
        private long _totalRead;

        internal LengthLimitedStream(Stream inner, long maxBytes, string description)
        {
            _inner = inner;
            _maxBytes = maxBytes;
            _description = description;
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            var read = _inner.Read(buffer, offset, count);
            Track(read);
            return read;
        }

        public override int Read(Span<byte> buffer)
        {
            var read = _inner.Read(buffer);
            Track(read);
            return read;
        }

        private void Track(int bytesJustRead)
        {
            if (bytesJustRead <= 0)
            {
                return;
            }

            _totalRead += bytesJustRead;
            if (_totalRead > _maxBytes)
            {
                // Plain integer interpolation (no ToString("N")/culture-sensitive thousands separator,
                // matching this repo's own invariant-formatting discipline elsewhere) plus
                // a friendly MB figure, so the message reads the same regardless of the host's culture.
                throw new ResourceLimitExceededException(
                    $"{_description} exceeds the supported size limit of {_maxBytes / (1024 * 1024)} MB "
                    + $"({_maxBytes} bytes); this looks like a malformed or maliciously crafted file rather "
                    + "than a real Business Central dataset/schema part.");
            }
        }

        public override bool CanRead => true;

        public override bool CanSeek => false;

        public override bool CanWrite => false;

        public override long Length => throw new NotSupportedException();

        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override void Flush() => _inner.Flush();

        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

        public override void SetLength(long value) => throw new NotSupportedException();

        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
