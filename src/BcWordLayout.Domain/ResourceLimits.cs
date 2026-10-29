using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

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
/// The OpenXml SDK's OWN part loads (<c>MainDocumentPart.Document</c>, headers/footers, settings, and every
/// part <c>OpenXmlValidator</c> walks) are recursive too, and are not this repo's code to bound from the
/// inside: measured on SDK 3.5.1, its <c>OpenXmlCompositeElement.Populate</c> overflows a default stack at
/// roughly 2,250 nested elements (a 1.4 KB <c>.docx</c>), and <c>OpenXmlValidator</c>'s package walk at
/// roughly 1,500 chained parts. So <see cref="EnsurePackageWithinLimits"/> bounds them from the OUTSIDE:
/// <see cref="LayoutPackage.Open"/> — the one way this codebase opens a <c>.docx</c> — scans every part's raw
/// XML with a non-recursive reader before the SDK parses any of it, enforcing
/// <see cref="MaxPackageXmlDepth"/>, <see cref="MaxPackageXmlPartBytes"/>, <see cref="MaxPackageXmlBytes"/>
/// and <see cref="MaxPackageParts"/>.
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
    /// Byte ceiling for the AL <c>app.json</c> <see cref="TargetVersionResolver"/> reads above a layout. A real
    /// one is a few KB; a larger file makes the target unknown, with the size in the reason, rather than being
    /// parsed.
    /// </summary>
    internal const long MaxAppJsonBytes = 1L * 1024 * 1024;

    /// <summary>
    /// Hard ceiling on the NUMBER of custom XML parts <see cref="SchemaProvider.FindBcPart"/>/
    /// <see cref="SchemaProvider.FindBcParts"/> will even attempt to enumerate. The per-part byte cap above
    /// bounds any ONE part's cost, but a package could still carry an enormous NUMBER of small,
    /// cheap-to-compress parts, making the per-part iteration/parse LOOP itself the unbounded cost - the byte
    /// cap alone does not bound it. A real BC layout carries a
    /// handful of custom XML parts (the BC dataset part, occasionally an unrelated Office bibliography part);
    /// 64 is generous headroom while still bounding the loop to a small, fast, always-survivable amount of
    /// work. Kept below <see cref="MaxPackageParts"/> so a package stuffed with custom XML parts trips THIS,
    /// more specific, check (whose message names custom XML parts) rather than the package-wide one.
    /// </summary>
    internal const int MaxCustomXmlParts = 64;

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
    /// Max element nesting depth of ANY XML part in a package, checked by
    /// <see cref="EnsurePackageWithinLimits"/> before the OpenXml SDK parses the part into its (recursive)
    /// DOM. The deepest part in the corpus nests 31 levels; the SDK's own load overflows the
    /// stack at roughly 2,250. 512 is 16x the corpus and about 4x under the crash, and stays above twice
    /// <see cref="MaxElementNestingDepth"/> (a content-control level is two XML levels, <c>w:sdt</c> plus
    /// <c>w:sdtContent</c>), so this repo's own walkers keep their own, more specific, depth error.
    /// </summary>
    internal const int MaxPackageXmlDepth = 512;

    /// <summary>
    /// Byte ceiling for any ONE XML part the OpenXml SDK may parse (<c>document.xml</c>, a header, styles,
    /// ...) — the package-wide counterpart of <see cref="MaxCustomXmlPartBytes"/>, which covers only the parts
    /// this repo loads itself. The largest part in the corpus is a 153 KB <c>document.xml</c>; 16 MB is a
    /// hundred times that, while bounding what a few-KB compressed "bomb" part can make the SDK allocate.
    /// </summary>
    internal const long MaxPackageXmlPartBytes = 16L * 1024 * 1024;

    /// <summary>
    /// Ceiling on the SUM of all XML part sizes in a package. <c>OpenXmlValidator</c> parses every part, so the
    /// per-part cap above alone would still let <see cref="MaxPackageParts"/> maximum-size parts add up to
    /// gigabytes. A real layout's XML totals well under 1 MB.
    /// </summary>
    internal const long MaxPackageXmlBytes = 64L * 1024 * 1024;

    /// <summary>
    /// Max number of distinct parts reachable in a package. Bounds every recursive walk over the part graph
    /// that this repo does NOT write — <c>OpenXmlValidator</c>'s package walk overflows the stack on a chain of
    /// roughly 1,500 parts — because no walk can nest deeper than the number of parts there are. Corpus
    /// layouts carry 18-31 zip entries; 256 is ample for a layout with many images while staying several
    /// times under that crash.
    /// </summary>
    internal const int MaxPackageParts = 256;

    /// <summary>
    /// Max nesting of Word fields (<c>w:fldChar</c> begin inside another field) that
    /// <c>BcWordLayout.Merge.ExternalReferenceNeutralizer</c> tracks while it reads a field code. Every
    /// element of a field code is attributed to each enclosing field still reading its own code, so the cost
    /// grows with nesting; real layouts nest two or three deep (an <c>IF</c> around <c>PAGE</c>/<c>NUMPAGES</c>).
    /// </summary>
    internal const int MaxFieldNestingDepth = 32;

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
    /// Checks <paramref name="package"/> against the package-wide caps (<see cref="MaxPackageParts"/>,
    /// <see cref="MaxPackageXmlPartBytes"/>, <see cref="MaxPackageXmlBytes"/>, <see cref="MaxPackageXmlDepth"/>)
    /// BEFORE the OpenXml SDK parses any of its parts, throwing <see cref="ResourceLimitExceededException"/> on
    /// the first one exceeded. Opening the package is safe to do first: measured, the SDK's open builds the
    /// part graph iteratively and parses no part content.
    /// <para>
    /// Both walks here are non-recursive (an explicit stack over the part graph, a streaming
    /// <see cref="XmlReader"/> over each part), so a crafted package cannot overflow THIS code's stack either.
    /// Every part is scanned whatever its content type claims: a binary part (an image) fails the reader on
    /// its first bytes and costs nothing.
    /// </para>
    /// <para>
    /// A part that is not well-formed XML simply ends its own scan at the error. That is safe because the
    /// scan's reader is never stricter than the SDK's (which prohibits DTDs, checks characters and requires a
    /// single root — measured): the SDK hits the same error at or before the same byte, so it never parses
    /// past the point this scan already cleared.
    /// </para>
    /// <para>
    /// Call it on a READ-ONLY open: in a writable package, System.IO.Packaging decompresses a whole part into
    /// memory the moment its stream is opened, before any cap here could apply. <see cref="LayoutPackage.Open"/>
    /// does exactly that for its writable callers.
    /// </para>
    /// </summary>
    internal static void EnsurePackageWithinLimits(OpenXmlPackage package)
    {
        var visited = new HashSet<OpenXmlPart>();
        var pending = new Stack<OpenXmlPart>(package.Parts.Select(p => p.OpenXmlPart));
        long totalXmlBytes = 0;

        while (pending.TryPop(out var part))
        {
            if (!visited.Add(part))
            {
                continue;
            }

            if (visited.Count > MaxPackageParts)
            {
                throw new ResourceLimitExceededException(
                    $"Package has more than {MaxPackageParts} parts, exceeding the supported limit; this looks "
                    + "like a malformed or maliciously crafted file rather than a real Business Central layout.");
            }

            totalXmlBytes += ScanPartXml(part);
            if (totalXmlBytes > MaxPackageXmlBytes)
            {
                throw new ResourceLimitExceededException(
                    $"Package XML parts total more than the supported size limit of {MaxPackageXmlBytes / (1024 * 1024)} "
                    + "MB; this looks like a malformed or maliciously crafted file rather than a real Business "
                    + "Central layout.");
            }

            foreach (var child in part.Parts)
            {
                pending.Push(child.OpenXmlPart);
            }
        }
    }

    /// <summary>
    /// Streams <paramref name="part"/> through a non-recursive reader, enforcing
    /// <see cref="MaxPackageXmlPartBytes"/> and <see cref="MaxPackageXmlDepth"/>, and returns the bytes read
    /// (for the <see cref="MaxPackageXmlBytes"/> total). See <see cref="EnsurePackageWithinLimits"/> for why a
    /// part that is not well-formed XML only ends its own scan.
    /// </summary>
    private static long ScanPartXml(OpenXmlPart part)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var limited = new LengthLimitedStream(stream, MaxPackageXmlPartBytes, $"Package part '{part.Uri}'");
        try
        {
            using var reader = XmlReader.Create(limited, CreateScanSettings());
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.Depth > MaxPackageXmlDepth)
                {
                    throw DepthExceeded($"XML part '{part.Uri}'", MaxPackageXmlDepth);
                }
            }
        }
        catch (XmlException)
        {
            // Not XML (a binary part) or malformed from here on: the SDK fails at or before this byte too.
        }

        return limited.BytesRead;
    }

    /// <summary>
    /// The scan's reader settings: deliberately MORE permissive than the SDK's own reader (DTDs ignored rather
    /// than prohibited, characters unchecked) so this scan can never stop on an error the SDK would read past.
    /// Nothing is resolved or expanded — an ignored DTD declares no entities.
    /// </summary>
    private static XmlReaderSettings CreateScanSettings() => new()
    {
        DtdProcessing = DtdProcessing.Ignore,
        XmlResolver = null,
        CheckCharacters = false,
        IgnoreComments = true,
        IgnoreWhitespace = true,
        IgnoreProcessingInstructions = true,
    };

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

        /// <summary>Bytes read through this wrapper so far.</summary>
        internal long BytesRead => _totalRead;

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
                    + "than a real Business Central layout, dataset or schema file.");
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
