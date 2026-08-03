using System.Globalization;
using System.Text;
using System.Xml.Linq;
using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;

namespace BcWordLayout.Merge;

/// <summary>
/// Generates a deterministic, type-aware sample dataset for a parsed BC report schema, or loads a real
/// exported dataset when <see cref="SampleDataOptions.DataOverridesPath"/> is supplied. The result is an
/// <see cref="XDocument"/> in the schema's own namespace, ready for the merge engine to resolve
/// <c>w:dataBinding</c> XPaths against. Building walks the <see cref="DatasetTree"/> model directly (no
/// XML re-parsing), so output shape always matches the schema exactly.
/// </summary>
public static class SampleDataGenerator
{
    private static readonly DateTime DateBase = new(2026, 1, 1);

    private static readonly string[] FillerWords =
    {
        "Alpha", "Beta", "Gamma", "Delta", "Sample", "Demo", "Preview", "Draft", "Example", "Test",
    };

    /// <summary>
    /// Short, deterministic marker substituted for an image-ish leaf column's value instead of fake text —
    /// see <see cref="GenerateLeafValue"/>'s remarks. Deliberately NOT the <c>«leaf?»</c> guillemet shape
    /// <see cref="MergeEngine"/> uses for an unresolved binding: that marker means "this
    /// binding failed to resolve" (a merge-time error), whereas this one means "this is a deliberately
    /// simplified mock value for a data type this generator does not fake" (never an error) — keeping the
    /// two visually distinct avoids a preview reader confusing one for the other.
    /// </summary>
    private const string ImageMarker = "[image]";

    /// <summary>
    /// Whole-word tokens (see <see cref="HasAnyWordToken"/>) that mark a leaf column as holding IMAGE data
    /// rather than text — see <see cref="GenerateLeafValue"/>'s remarks.
    /// </summary>
    private static readonly string[] ImageColumnTokens = { "Picture", "Image", "Logo", "Photo", "Bitmap" };

    /// <summary>
    /// Builds a sample dataset for <paramref name="schema"/>. When <see cref="SampleDataOptions.DataOverridesPath"/>
    /// is set, that file is loaded and validated in place of generation (see the type doc remarks); otherwise
    /// the tree is walked and every leaf column receives a seeded, type-aware fake value.
    /// </summary>
    public static SampleDataset Generate(DatasetTree schema, SampleDataOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(schema);
        options ??= new SampleDataOptions();

        if (!string.IsNullOrEmpty(options.DataOverridesPath))
        {
            return LoadOverrides(options.DataOverridesPath, schema.Report);
        }

        var ns = XNamespace.Get(schema.Report.Namespace);
        var ctx = new GenerationContext
        {
            Ns = ns,
            Rng = new Random(options.Seed),
            Rows = options.Rows,
            Cap = options.MaxRowsPerItem,
            RemainingBudget = Math.Max(0, options.MaxTotalInstances),
            RepeaterConsumedPaths = options.RepeaterConsumedPaths,
        };

        // The root element itself is the document container, not a business instance, so it is never counted
        // against RemainingBudget - only its descendant instances are.
        var root = BuildInstance(schema.Root, ctx, instanceIndex: 0);

        return new SampleDataset
        {
            Xml = new XDocument(root),
            Namespace = schema.Report.Namespace,
            Truncated = ctx.Truncated,
        };
    }

    // ---- generation: schema tree walk ----

    /// <summary>
    /// Mutable state threaded through the recursive <see cref="BuildInstance"/> walk: the fixed generation
    /// inputs plus the shared global instance budget (<see cref="RemainingBudget"/> / <see cref="Truncated"/>)
    /// that <see cref="SampleDataOptions.MaxTotalInstances"/> enforces across the WHOLE tree (a per-item cap
    /// cannot, because instance counts multiply across nesting depth).
    /// </summary>
    private sealed class GenerationContext
    {
        public required XNamespace Ns { get; init; }
        public required Random Rng { get; init; }
        public required int Rows { get; init; }
        public required int Cap { get; init; }
        public int RemainingBudget { get; set; }
        public bool Truncated { get; set; }

        /// <summary>See <see cref="SampleDataOptions.RepeaterConsumedPaths"/>; null means "treat every business item as consumed".</summary>
        public IReadOnlySet<string>? RepeaterConsumedPaths { get; init; }
    }

    /// <summary>
    /// Builds one instance of <paramref name="item"/>: its leaf columns (in document order) followed by
    /// a fresh, independently generated set of child-item instances. System subtrees (<see cref="DataItem.IsSystem"/>)
    /// always get exactly 1 instance; a business item gets <c>Math.Min(Rows, Cap)</c> instances when it is
    /// REPEATER-CONSUMED (see <see cref="IsRepeaterConsumed"/>/<see cref="SampleDataOptions.RepeaterConsumedPaths"/>)
    /// and exactly 1 otherwise (see <see cref="SampleDataOptions.MaxRowsPerItem"/>'s own remarks
    /// for why generation additionally bounds a consumed item's own count this way) — each parent instance owns
    /// its own recursively generated children, so e.g. a repeater-consumed Line under a Header yields that many
    /// Lines per Header, not that many total.
    /// <para>
    /// Every business (non-system) instance is charged against the shared
    /// <see cref="GenerationContext.RemainingBudget"/> as it is created (depth-first, document order); once the
    /// budget is exhausted no further business instances are generated and
    /// <see cref="GenerationContext.Truncated"/> is set. This is the only thing that bounds a deeply-nested
    /// schema's <c>count^depth</c> blow-up; the per-item cap alone cannot (see
    /// <see cref="SampleDataOptions.MaxTotalInstances"/>). System instances (always exactly 1) are never charged.
    /// Capping every UNCONSUMED item to 1 instance (rather than <c>Math.Min(Rows, Cap)</c>) means an unconsumed
    /// subtree can no longer eat this budget disproportionately either — the whole point: a
    /// deeply nested schema no longer burns the shared budget generating rows for data items the document
    /// never actually repeats, at the expense of the ones it does.
    /// </para>
    /// <para>
    /// SECURITY: this method recurses once per <see cref="DataItem"/>
    /// nesting level, with no depth counter of its own — deliberately: every <see cref="DataItem"/> tree is
    /// built by <see cref="SchemaProvider.BuildNode"/>, the ONE place in the process that constructs one, and
    /// that method already rejects (via <see cref="ResourceLimits.MaxSchemaDepth"/>) any tree deeper than the
    /// cap before it ever reaches here. A crafted schema with pathological nesting therefore never produces a
    /// <see cref="DataItem"/> tree for this method to walk in the first place; a second counter here would be
    /// redundant.
    /// </para>
    /// </summary>
    private static XElement BuildInstance(DataItem item, GenerationContext ctx, int instanceIndex)
    {
        var element = new XElement(ctx.Ns + item.Name);

        foreach (var column in item.Columns)
        {
            element.Add(new XElement(ctx.Ns + column.Name, GenerateLeafValue(column, ctx.Rng, instanceIndex)));
        }

        foreach (var child in item.Children)
        {
            var count = child.IsSystem
                ? 1
                : IsRepeaterConsumed(child, ctx)
                    ? Math.Min(ctx.Rows, ctx.Cap)
                    : Math.Min(1, ctx.Cap);

            for (var i = 0; i < count; i++)
            {
                if (!child.IsSystem)
                {
                    if (ctx.RemainingBudget <= 0)
                    {
                        ctx.Truncated = true;
                        break;
                    }

                    ctx.RemainingBudget--;
                }

                element.Add(BuildInstance(child, ctx, i));
            }
        }

        return element;
    }

    /// <summary>
    /// True when <paramref name="item"/> should be multiplied to <c>Math.Min(Rows, Cap)</c> instances rather
    /// than capped to exactly 1. <see cref="GenerationContext.RepeaterConsumedPaths"/> null
    /// (no live document was scanned — every direct <see cref="Generate"/> caller that hands a bare schema,
    /// e.g. every existing unit test) means "treat everything as consumed", preserving the earlier behavior
    /// exactly; otherwise an item is consumed only when some repeating section in the document is actually
    /// bound to its own <see cref="DataItem.Path"/> (see <c>MergeEngine.ScanRepeaterConsumedPaths</c>).
    /// </summary>
    private static bool IsRepeaterConsumed(DataItem item, GenerationContext ctx) =>
        ctx.RepeaterConsumedPaths is null || ctx.RepeaterConsumedPaths.Contains(item.Path);

    // ---- leaf value generation (type inference from the column name) ----

    /// <summary>
    /// Picks a type-aware fake value for a single leaf column. Label columns (<see cref="DatasetColumn.IsLabel"/>)
    /// get a static humanized caption; everything else is inferred from whole camelCase/underscore-split
    /// words of the column name (never a raw substring match — see <see cref="HasAnyWordToken"/>) and varies
    /// with <paramref name="rng"/> and <paramref name="instanceIndex"/> only, so sibling rows look distinct
    /// while staying fully deterministic.
    /// </summary>
    /// <remarks>
    /// Known fidelity gap: a leaf column whose name looks like it holds IMAGE data (real corpus
    /// examples: <c>CompanyPicture</c>/<c>CompanyInfo1Picture</c>/<c>CompanyInfo2Picture</c> — present in
    /// the schema but, unlike <c>CompanyPicture</c> itself, never bound by any control in those layouts)
    /// gets a short <see cref="ImageMarker"/> instead of a fake value here. BC's real
    /// dataset stores such a column as (potentially large) base64 image bytes; a plain-text FIELD control
    /// bound to one — hand-authored, or produced by <c>insert_field</c>, which has no way to know a leaf
    /// column "is really" image data — would otherwise show either lorem-ish fallback text (this generator,
    /// unpatched) or a giant base64 blob (a real <see cref="SampleDataOptions.DataOverridesPath"/> export) in
    /// a preview. BC itself renders the actual image for such a binding; this mock cannot (only a dedicated
    /// PICTURE control — <c>w:picture</c>, filled by <see cref="MergeEngine"/> with the
    /// placeholder PNG — gets a real image), so the short marker is the honest substitute.
    /// </remarks>
    private static string GenerateLeafValue(DatasetColumn column, Random rng, int instanceIndex)
    {
        if (column.IsLabel)
        {
            return HumanizeLabel(column.Name);
        }

        var name = column.Name;
        var words = Tokenize(name);

        // Caption-indicator words trump every TYPE token below, deliberately checked first (2026-07-31
        // preview sweep): BC's standard layouts bind whole header rows to columns like
        // DescCustLedgEntry2Caption / DueDateCaption and totals-row labels to TotalText-style columns.
        // Type-first inference turned those into dates, DOCN- codes, and bare decimals sitting where
        // column headings belong ("Total" cells showing 91.93 next to the actual amount). A name that
        // says "I hold display text ABOUT a value" must never inherit that value's own type. This is a
        // sample-value decision only — it does not classify the column as a label (IsLabel above), so
        // control classification and label counts are untouched. Checked before the image tokens too:
        // a CompanyPictureCaption holds the picture's caption text, not image bytes.
        if (HasAnyWordToken(words, CaptionIndicatorTokens))
        {
            return HumanizeCaptionName(words);
        }

        if (HasAnyWordToken(words, ImageColumnTokens))
        {
            return ImageMarker;
        }

        // RULE ORDER BELOW IS LOAD-BEARING (2026-07-31 preview sweep). The specific orderings that fix
        // real corpus misfires:
        //   boolean before amount   - PricesInclVAT is a Yes/No flag whose name contains the amount
        //                             token "VAT"; amount-first rendered it as 72.26.
        //   No/Code before amount   - CompanyVATRegistrationNo is an identifier whose name contains
        //                             "VAT"; amount-first rendered it as 176.94.
        //   No/Code AFTER phone     - PhoneNo must keep its phone-shaped sample, not become PHON-0042.
        //   percent before amount   - VATPct_Line / *DiscountPercent carry amount tokens; amount-first
        //                             sampled them 0-1000 (925.92 "percent"), percent clamps to 0-100.

        if (HasAnyWordToken(words, "Is", "Has", "Show", "Print")
            || (HasAnyWordToken(words, "Incl", "Including") && HasAnyWordToken(words, "Prices")))
        {
            return GenerateYesNo(rng, instanceIndex);
        }

        // Unit-of-measure columns occupy the narrowest cell of every corpus line table; the generic
        // fallback ("Unit Of Measure Preview 1", ~25 chars) wrapped letter-by-letter there. Real BC data
        // is a short code (PCS, BOX, ...), so the sample is too. "Unit" alone must NOT match: UnitPrice /
        // UnitCost are amounts.
        if (HasAnyWordToken(words, "UOM") || (HasAnyWordToken(words, "Unit") && HasAnyWordToken(words, "Measure")))
        {
            return GenerateUnitOfMeasure(rng, instanceIndex);
        }

        if (HasAnyWordToken(words, "Email"))
        {
            return GenerateEmail(words, rng, instanceIndex);
        }

        if (HasAnyWordToken(words, "Phone"))
        {
            return GeneratePhone(rng, instanceIndex);
        }

        // "Identifier" rides with No/Code (and shares their before-amount precedence): VATIdentifier in
        // the corpus VAT-clause tables is a code like "VAT25", not the 125.25 the VAT token made of it.
        if (HasAnyWordToken(words, "No", "Code", "Identifier"))
        {
            return GenerateCode(name, rng, instanceIndex);
        }

        if (HasAnyWordToken(words, "Pct", "Percent", "Percentage"))
        {
            return GeneratePercent(rng, instanceIndex);
        }

        // "Dt" and "Today" are corpus-proven date spellings the plain "Date" token missed:
        // StandardPurchaseOrder's ExptRcptDt_PurchHeader rendered as name-echo text where a date
        // belongs, and StandardStatement's TodayFormatted is BC's pre-formatted "today" string.
        if (HasAnyWordToken(words, "Date", "Dt", "Today"))
        {
            return GenerateDate(rng, instanceIndex);
        }

        // "Amt" carries the same weight as "Amount": BC dataset columns abbreviate aggressively
        // (OriginalAmt_CustLedgEntry2, RemainAmt_..., AgingBandBufCol1Amt_... across StandardStatement),
        // and before it was listed every money column in that layout fell through to name-echo TEXT —
        // whole amount columns reading "Remain Amt Cust Ledg Entries Delta 1" in a preview. "Balance" is
        // the same story (CustBalance_*, StartBalance_*, OverdueBalance_* — all money), and so is the
        // abbreviation "Disc" (LineDisc_PurchLine, InvDiscAmt_PurchLine in StandardPurchaseOrder).
        if (HasAnyWordToken(words, "Amount", "Amt", "Balance", "Price", "Cost", "Total", "Discount", "Disc", "VAT", "Tax"))
        {
            return GenerateDecimal(rng, instanceIndex);
        }

        // "Days" (ContractBillingDetailsDays in SalesInvoiceForSubscriptionBilling) is an integer count;
        // the quantity range reads naturally for it.
        if (HasAnyWordToken(words, "Qty", "Quantity", "Days"))
        {
            return GenerateQuantity(rng, instanceIndex);
        }

        return GenerateFallbackText(name, rng, instanceIndex);
    }

    private static string GenerateDate(Random rng, int instanceIndex)
    {
        var offset = rng.Next(0, 365) + instanceIndex;
        return DateBase.AddDays(offset).ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
    }

    private static string GenerateDecimal(Random rng, int instanceIndex)
    {
        var hundredths = rng.Next(0, 100_000);
        var value = (hundredths / 100m) + (instanceIndex * 10m);
        return value.ToString("0.00", CultureInfo.InvariantCulture);
    }

    private static string GenerateQuantity(Random rng, int instanceIndex)
    {
        var qty = rng.Next(1, 50) + instanceIndex;
        return qty.ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Realistic short unit-of-measure codes; see the UOM rule in <see cref="GenerateLeafValue"/> for why
    /// these are not fallback text.
    /// </summary>
    private static readonly string[] UnitOfMeasureCodes = ["PCS", "BOX", "SET", "KG", "HRS"];

    private static string GenerateUnitOfMeasure(Random rng, int instanceIndex)
    {
        return UnitOfMeasureCodes[(rng.Next(UnitOfMeasureCodes.Length) + instanceIndex) % UnitOfMeasureCodes.Length];
    }

    /// <summary>
    /// Percent sample bounded to a plausible 0.00–100.00 — never the unconstrained decimal range, which
    /// produced "VAT Pct 925.92" in previews. Two decimals to match how BC formats percentages.
    /// </summary>
    private static string GeneratePercent(Random rng, int instanceIndex)
    {
        var hundredths = (rng.Next(0, 10_001) + (instanceIndex * 250)) % 10_001;
        return (hundredths / 100m).ToString("0.00", CultureInfo.InvariantCulture);
    }

    /// <summary>Boolean flag sample, in the "Yes"/"No" spelling BC's report datasets use for booleans.</summary>
    private static string GenerateYesNo(Random rng, int instanceIndex)
    {
        return (rng.Next(0, 2) + instanceIndex) % 2 == 0 ? "Yes" : "No";
    }

    private static string GenerateEmail(string[] words, Random rng, int instanceIndex)
    {
        var local = (words.Length > 0 ? words[0] : "user").ToLowerInvariant();
        var suffix = rng.Next(10, 99);
        var instanceText = (instanceIndex + 1).ToString(CultureInfo.InvariantCulture);
        var suffixText = suffix.ToString(CultureInfo.InvariantCulture);
        return $"{local}{instanceText}{suffixText}@example.com";
    }

    private static string GeneratePhone(Random rng, int instanceIndex)
    {
        var exchange = rng.Next(200, 999);
        var line = rng.Next(1000, 9999);
        var exchangeText = ((exchange + instanceIndex) % 1000).ToString("000", CultureInfo.InvariantCulture);
        var lineText = line.ToString("0000", CultureInfo.InvariantCulture);
        return $"+1-555-{exchangeText}-{lineText}";
    }

    private static string GenerateCode(string name, Random rng, int instanceIndex)
    {
        var letters = new string(name.Where(char.IsLetter).ToArray()).ToUpperInvariant();
        var prefix = letters.Length switch
        {
            0 => "CODE",
            <= 4 => letters,
            _ => letters[..4],
        };

        var number = rng.Next(1, 9999) + instanceIndex;
        var numberText = number.ToString("0000", CultureInfo.InvariantCulture);
        return $"{prefix}-{numberText}";
    }

    private static string GenerateFallbackText(string name, Random rng, int instanceIndex)
    {
        var humanized = HumanizeWords(name);
        var filler = FillerWords[rng.Next(FillerWords.Length)];
        var instanceText = (instanceIndex + 1).ToString(CultureInfo.InvariantCulture);
        return string.IsNullOrEmpty(humanized)
            ? $"{filler} {instanceText}"
            : $"{humanized} {filler} {instanceText}";
    }

    // ---- naming helpers ----

    /// <summary>
    /// Whole-word tokens that mark a column as holding DISPLAY TEXT about a value rather than the value
    /// itself (<c>DueDateCaption</c>, <c>TotalText</c>, <c>*_Lbl</c> mid-name). Matched before every type
    /// token in <see cref="GenerateLeafValue"/> — see the comment at that call site.
    /// </summary>
    private static readonly string[] CaptionIndicatorTokens = ["Caption", "Label", "Lbl", "Text"];

    /// <summary>
    /// Caption-style sample for a caption-indicator column: the humanized name minus the indicator words
    /// themselves (<c>DueDateCaption</c> → <c>Due Date</c>, <c>TotalText</c> → <c>Total</c>). No filler /
    /// per-row variance on purpose — a real caption is constant across rows, and a header cell reading
    /// "Due Date" is the closest a mock render gets to the real BC caption. Falls back to the full
    /// humanized name when the indicator words are all there is (e.g. a column literally named
    /// <c>Text</c>).
    /// </summary>
    private static string HumanizeCaptionName(IReadOnlyList<string> words)
    {
        var kept = words
            .Where(w => !CaptionIndicatorTokens.Any(t => string.Equals(w, t, StringComparison.OrdinalIgnoreCase)))
            .ToList();
        return string.Join(' ', kept.Count == 0 ? words : kept);
    }

    /// <summary>
    /// Strips the active label convention's suffix (see <see cref="LabelConvention.StripSuffix"/> — by
    /// default <c>_Lbl</c>/<c>Lbl</c>), then one trailing caption-indicator token the convention itself
    /// does not know about, then humanizes the remainder. The second strip matters for labels classified
    /// by the labels-data-item rule rather than by suffix (see <see cref="LabelConvention"/>'s remarks):
    /// their names end in <c>Caption</c>/<c>Label</c> (<c>SalesLineDocumentNoCaption</c>,
    /// <c>ShipmentDateLabel</c>), and a header cell reading "Sales Line Document No" is the real-caption
    /// mock — "Sales Line Document No Caption" is name-echo noise. Trailing-only and single-pass on
    /// purpose (mirroring <see cref="LabelConvention.StripSuffix"/>), so a mid-name indicator word is
    /// never touched and a column named exactly one indicator word keeps its whole name.
    /// </summary>
    private static string HumanizeLabel(string columnName)
    {
        var stripped = StripTrailingCaptionIndicator(LabelConvention.Current.StripSuffix(columnName));
        var humanized = HumanizeWords(stripped);
        return humanized.Length == 0 ? stripped : humanized;
    }

    /// <summary>
    /// Removes ONE trailing <see cref="CaptionIndicatorTokens"/> token (longest first, ordinal), plus a
    /// joining underscore if one remains — <c>No_ItemCaption</c> → <c>No_Item</c>. Returns
    /// <paramref name="name"/> unchanged when no token matches or the token is all there is.
    /// </summary>
    private static string StripTrailingCaptionIndicator(string name)
    {
        foreach (var token in CaptionIndicatorTokens.OrderByDescending(t => t.Length))
        {
            if (name.Length > token.Length && name.EndsWith(token, StringComparison.Ordinal))
            {
                var trimmed = name[..^token.Length];
                return trimmed.EndsWith('_') ? trimmed[..^1] : trimmed;
            }
        }

        return name;
    }

    /// <summary>
    /// Inserts spaces at camelCase / letter-digit boundaries and turns underscores into spaces, e.g.
    /// <c>TotalAmount</c> -&gt; <c>Total Amount</c>, <c>BilledTo</c> -&gt; <c>Billed To</c>.
    /// </summary>
    private static string HumanizeWords(string name)
    {
        var sb = new StringBuilder(name.Length + 8);

        for (var i = 0; i < name.Length; i++)
        {
            var c = name[i];
            if (c == '_')
            {
                sb.Append(' ');
                continue;
            }

            if (i > 0)
            {
                var prev = name[i - 1];
                var boundary =
                    (char.IsLower(prev) && char.IsUpper(c)) ||
                    (char.IsUpper(prev) && char.IsUpper(c) && i + 1 < name.Length && char.IsLower(name[i + 1])) ||
                    (char.IsLetter(prev) && char.IsDigit(c)) ||
                    (char.IsDigit(prev) && char.IsLetter(c));

                if (boundary)
                {
                    sb.Append(' ');
                }
            }

            sb.Append(c);
        }

        var words = sb.ToString().Split(' ', StringSplitOptions.RemoveEmptyEntries);
        return string.Join(' ', words);
    }

    /// <summary>Splits a column name into its camelCase/underscore words, for whole-word type inference.</summary>
    private static string[] Tokenize(string name) =>
        HumanizeWords(name).Split(' ', StringSplitOptions.RemoveEmptyEntries);

    /// <summary>
    /// True when any of <paramref name="tokens"/> equals (case-insensitive) a whole word in <paramref name="words"/>.
    /// Whole-word matching — never a raw substring <c>Contains</c> — is deliberate: it avoids false positives
    /// such as "LastUpdatedBy" (contains "date"), "PrivateNote" / "Innovation" (contain "vat").
    /// </summary>
    private static bool HasAnyWordToken(IReadOnlyList<string> words, params string[] tokens) =>
        tokens.Any(t => words.Any(w => string.Equals(w, t, StringComparison.OrdinalIgnoreCase)));

    // ---- data overrides ----

    /// <summary>
    /// Loads a real exported BC dataset XML as the sample dataset, accepting BOTH encodings BC produces
    /// (sniffed by root element): the layout's own data-store part shape (root
    /// <see cref="OoxmlNames.RootElementName"/>, namespace starting <see cref="OoxmlNames.BcNamespacePrefix"/>,
    /// validated loosely and returned verbatim) and the report UI's *Send to → XML* export (root
    /// <see cref="ReportDataSetConverter.ExportRootElementName"/>, converted into the first shape in
    /// <paramref name="report"/>'s namespace — see <see cref="ReportDataSetConverter"/>, including per-column
    /// <c>decimalformatter</c> application). Mirrors <see cref="SchemaProvider.FromSchemaXml"/>:
    /// <see cref="XDocument.Load(System.IO.Stream)"/> reads the encoding declaration from the stream itself,
    /// so UTF-16 LE + BOM exports work unchanged.
    /// </summary>
    private static SampleDataset LoadOverrides(string path, ReportIdentity report)
    {
        if (!File.Exists(path))
        {
            throw new FileNotFoundException("Data overrides file not found.", path);
        }

        XDocument xdoc;
        using (var stream = File.OpenRead(path))
        {
            // ResourceLimits.LoadXDocumentCapped enforces the size cap via a length-
            // limiting wrapper before handing the stream to the XmlReader; otherwise behaves exactly like
            // the bare XDocument.Load(stream) this replaces (same encoding-declaration sniffing).
            xdoc = ResourceLimits.LoadXDocumentCapped(stream, $"Data overrides file '{path}'");
        }

        var root = xdoc.Root
            ?? throw new InvalidDataException("Data overrides XML has no root element.");

        if (string.Equals(root.Name.LocalName, ReportDataSetConverter.ExportRootElementName, StringComparison.Ordinal))
        {
            xdoc = ReportDataSetConverter.ToNavWordReportXmlPart(xdoc, report);
            root = xdoc.Root!;
        }
        else if (!string.Equals(root.Name.LocalName, OoxmlNames.RootElementName, StringComparison.Ordinal))
        {
            throw new InvalidDataException(
                $"Data overrides root element is '{root.Name.LocalName}', expected "
                + $"'{OoxmlNames.RootElementName}' (a layout's data-store part) or "
                + $"'{ReportDataSetConverter.ExportRootElementName}' (the report UI's Send to → XML export).");
        }

        if (!root.Name.NamespaceName.StartsWith(OoxmlNames.BcNamespacePrefix, StringComparison.Ordinal))
        {
            throw new InvalidDataException(
                $"Data overrides namespace '{root.Name.NamespaceName}' does not start with "
                + $"'{OoxmlNames.BcNamespacePrefix}'.");
        }

        return new SampleDataset { Xml = xdoc, Namespace = root.Name.NamespaceName };
    }
}
