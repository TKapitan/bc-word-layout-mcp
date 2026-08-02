using BcWordLayout.Domain.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Ds = DocumentFormat.OpenXml.CustomXmlDataProperties;

namespace BcWordLayout.Domain;

/// <summary>
/// Creates a brand-new BC Word report layout <c>.docx</c> from a report dataset schema — either an existing
/// layout's own BC custom XML part (raw bytes copied verbatim) or a standalone exported schema <c>.xml</c>
/// — optionally starting from an UNBOUND branded/styled <c>.docx</c> template. The result is immediately
/// editable by <see cref="LayoutEditor"/>'s insert_* operations: it carries exactly one BC dataset custom XML
/// part (with a fresh <c>ds:itemID</c>), a glossary part with the <c>DefaultPlaceholder_-1854013440</c>
/// docPart entry every control <see cref="SdtFactory"/> builds references, a valid document body, and — for a
/// BLANK build — the empty header/footer parts a <c>layoutPart='header'/'footer'</c> insert needs to have
/// something to resolve against (see <see cref="HeaderFooterScaffold"/>; a template keeps its own). The
/// build is written atomically (assembled in a temp file next to the destination, validated, then moved into
/// place) and its own <see cref="LayoutValidator.Quick"/> result travels with it on
/// <see cref="CreateResult.QuickValidation"/> — a <c>templatePath</c> whose body already carried bound
/// content controls of its own that go stale once its BC part is replaced is refused outright (see
/// <see cref="TemplateNotUnboundException"/>) rather than reported as data on an otherwise successful result.
/// </summary>
public static class LayoutBuilder
{
    /// <summary>
    /// Creates <paramref name="outputPath"/> from <paramref name="schemaSource"/>, optionally copying
    /// <paramref name="templatePath"/> first and attaching the BC part to that copy instead of a blank
    /// document.
    /// </summary>
    /// <param name="schemaSource">
    /// Either an absolute path to an existing <c>.docx</c> layout (its BC dataset custom XML part is located
    /// via <see cref="SchemaProvider.FindBcPart"/> and its raw bytes are copied byte-for-byte — real corpus
    /// parts are UTF-16 LE with a BOM, and this preserves whatever encoding the source actually used) or a
    /// standalone exported schema <c>.xml</c> (validated via <see cref="SchemaProvider.FromSchemaXml"/>; its
    /// raw bytes are used as-is).
    /// </param>
    /// <param name="outputPath">
    /// Absolute path to write the new layout to. Overwritten if it already exists; its parent directory is
    /// created if missing.
    /// </param>
    /// <param name="templatePath">
    /// Optional absolute path to a <c>.docx</c> to start from instead of a blank document. This should be an
    /// UNBOUND branded/styled shell — headers/footers, a logo, fonts/styles — not a full BC layout with its
    /// own bound content controls: the template is copied first and its own body is kept as-is (a heading is
    /// appended only when the body is entirely empty — see <see cref="EnsureBodyHasContent"/>), but any
    /// EXISTING bound controls in that body are never stripped or rebound. When the template already has its
    /// own BC dataset custom XML part, it is removed and replaced by the one built from
    /// <paramref name="schemaSource"/> (see <see cref="CreateResult.ReplacedExistingBcPart"/>) — if the
    /// template also had bound controls of its own, they were built against ITS ORIGINAL schema/storeItemID
    /// and would go stale relative to the fresh one just attached; see <see cref="TemplateNotUnboundException"/>,
    /// which this method throws instead of silently reporting a "successful" result that ships that damage.
    /// Any unrelated custom XML part (e.g. the Office bibliography part some corpus layouts carry) is left
    /// untouched.
    /// </param>
    /// <exception cref="FileNotFoundException">
    /// <paramref name="schemaSource"/> or <paramref name="templatePath"/> does not exist.
    /// </exception>
    /// <exception cref="InvalidDataException">
    /// <paramref name="schemaSource"/> is a <c>.docx</c> with no main document part or no BC dataset custom
    /// XML part; is a schema <c>.xml</c> whose root is not <c>NavWordReportXmlPart</c>; or
    /// <paramref name="templatePath"/> has no main document part.
    /// </exception>
    /// <exception cref="TemplateNotUnboundException">
    /// <paramref name="templatePath"/> already carried its own BC dataset custom XML part AND bound content
    /// controls that go stale once that part is replaced by the one built from <paramref name="schemaSource"/>
    /// — see that type's own remarks for why this refuses rather than reports a warning.
    /// <paramref name="outputPath"/> is never touched when this is thrown.
    /// </exception>
    /// <exception cref="FileFormatException">
    /// The assembled layout failed <see cref="OpenXmlValidator"/> — an internal bug, not expected to be
    /// reachable in practice (see <see cref="AttachEverything"/>'s remarks). <paramref name="outputPath"/> is
    /// never touched when this is thrown.
    /// </exception>
    public static CreateResult Create(string schemaSource, string outputPath, string? templatePath = null, string? headingText = null)
    {
        if (string.IsNullOrWhiteSpace(schemaSource))
        {
            throw new ArgumentException("Schema source path must not be empty.", nameof(schemaSource));
        }

        if (string.IsNullOrWhiteSpace(outputPath))
        {
            throw new ArgumentException("Output path must not be empty.", nameof(outputPath));
        }

        if (!File.Exists(schemaSource))
        {
            throw new FileNotFoundException("schemaSource does not point to an existing file.", schemaSource);
        }

        if (templatePath is not null && !File.Exists(templatePath))
        {
            throw new FileNotFoundException("templatePath does not point to an existing file.", templatePath);
        }

        // Read schemaSource to completion into memory BEFORE outputPath (or the temp file below) is ever
        // touched - this, combined with never opening outputPath itself until the final atomic move, is what
        // makes schemaSource/templatePath/outputPath all safely aliasable (e.g. outputPath == templatePath,
        // "refresh this layout from itself").
        var (datasetBytes, identity) = LoadDatasetSource(schemaSource);

        // Path.GetDirectoryName returns null only for a bare root (e.g. "C:\") - an edge case in path
        // shape, not "this path does not exist in the layout"; left as a plain InvalidOperationException
        // (→ internal_error), not NotFoundException.
        var fullOutputPath = Path.GetFullPath(outputPath);
        var targetDir = Path.GetDirectoryName(fullOutputPath)
            ?? throw new InvalidOperationException($"Could not determine the directory of '{outputPath}'.");
        Directory.CreateDirectory(targetDir);

        var usedTemplate = templatePath is not null;

        // Build into a private temp file in outputPath's OWN directory (guaranteeing the commit below is a
        // same-volume rename) rather than outputPath itself, so a failure at any point - a template with no
        // main document part, schemaSource/templatePath/outputPath aliasing the same file, an (unreachable in
        // practice) internal bug producing structurally invalid OOXML - leaves outputPath completely
        // untouched instead of partially written or corrupted.
        var tempPath = Path.Combine(targetDir, $".bcwl-build-{Guid.NewGuid():N}.docx");
        try
        {
            string storeItemId;
            bool replacedExistingBcPart;
            ValidationResult quick;

            if (usedTemplate)
            {
                TransientFileRetry.Run(() => File.Copy(templatePath!, tempPath, overwrite: true));

                using var doc = WordprocessingDocument.Open(tempPath, true);
                _ = doc.MainDocumentPart
                    ?? throw new InvalidDataException($"Template '{templatePath}' has no main document part.");

                (storeItemId, replacedExistingBcPart, quick) =
                    AttachEverything(doc, datasetBytes, identity, headingText, scaffoldHeaderFooter: false);
            }
            else
            {
                using var doc = WordprocessingDocument.Create(tempPath, WordprocessingDocumentType.Document);
                var main = doc.AddMainDocumentPart();
                main.Document = new Document(new Body());

                (storeItemId, replacedExistingBcPart, quick) =
                    AttachEverything(doc, datasetBytes, identity, headingText, scaffoldHeaderFooter: true);
            }

            // REFUSE rather than half-succeed: a template that already carried its
            // own BC dataset part AND bound content controls of its own goes stale the moment that part is
            // replaced - every one of those controls' storeItemID now names a part this build just deleted,
            // and there is no honest way to fix that in place here (see TemplateNotUnboundException's own
            // remarks for why rebinding/stripping automatically would be guessing at caller intent, not
            // fixing a bug). Thrown BEFORE outputPath is ever touched - the temp file this build assembled
            // into is cleaned up by the finally block below exactly like any other Create-time failure, and
            // outputPath itself (whatever it held before, if anything) is left completely untouched.
            //
            // Keyed on ACTUAL staleness, not the part's mere presence: a template whose BC part had zero
            // bound controls still succeeds below - see CreateResult.ReplacedExistingBcPart, which stays
            // true for that case. (Zero controls is the ONLY replace-and-succeed shape - the fresh part
            // always carries a newly minted storeItemID, so any surviving bound control fails the
            // store-item-id check regardless of schema similarity.)
            if (replacedExistingBcPart && quick.ErrorCount > 0)
            {
                throw new TemplateNotUnboundException(
                    $"templatePath '{templatePath}' is a full BC layout, not an unbound shell: it already "
                    + "carries its own BC dataset custom XML part AND bound content controls built against "
                    + $"that part's original schema/storeItemID. Replacing the part with the one built from "
                    + "schemaSource (which create_layout always does) would leave those controls stale "
                    + $"({quick.ErrorCount} post-build quick-validation errors, mostly attributable to the "
                    + "stale bindings). Nothing was written.",
                    quick.ErrorCount);
            }

            // Only now - once the built layout has passed OpenXmlValidator (inside AttachEverything) and the
            // stale-controls refusal above did not fire - does outputPath itself get touched, and via a
            // same-volume rename rather than a truncate-and-stream copy, so it can never end up partially
            // written.
            TransientFileRetry.Run(() => File.Move(tempPath, fullOutputPath, overwrite: true));

            return new CreateResult
            {
                OutputPath = fullOutputPath,
                ReportName = identity.ReportName,
                ReportId = identity.ReportId,
                Namespace = identity.Namespace,
                StoreItemId = storeItemId,
                UsedTemplate = usedTemplate,
                ReplacedExistingBcPart = replacedExistingBcPart,
                QuickValidation = quick,
            };
        }
        finally
        {
            if (File.Exists(tempPath))
            {
                File.Delete(tempPath);
            }
        }
    }

    // ---- shared attach sequence (fresh document or template copy converge here) ----

    /// <summary>
    /// Attaches the BC part/glossary/body to <paramref name="doc"/> (a temp-file package, not yet committed
    /// to the caller's real output path), then runs <see cref="OpenXmlValidator"/> against it as a hard gate
    /// — <see cref="SdtFactory"/> and this type are both engineered to only ever emit/leave valid OOXML, so a
    /// structural error here would mean an internal bug; rather than write a broken file,
    /// <see cref="LayoutBuilder.Create"/> throws <see cref="FileFormatException"/> and leaves the caller's
    /// output path untouched (this path is not expected to be reachable in practice and is deliberately not
    /// force-tested — there is no honest way to trigger it from the public surface without fabricating an
    /// internal bug to order, mirroring <c>BcWordLayout.McpHost.Tools.ToolGuards.GuardEdit</c>'s own
    /// analogous backstop). <paramref name="scaffoldHeaderFooter"/> is true only for the BLANK build path —
    /// a <c>templatePath</c> brings its own header/footer story (branded letterhead, a distinct first page)
    /// and must never have parts injected into it. <see cref="LayoutValidator.Quick"/> is then run and returned as-is — unlike the
    /// hard OpenXmlValidator gate, a non-zero <see cref="ValidationResult.ErrorCount"/> here does NOT throw
    /// FROM THIS METHOD: the decision to refuse when a template's own pre-existing bound controls go stale
    /// against the freshly attached BC part is made by the caller (<see cref="LayoutBuilder.Create"/>, which
    /// throws <see cref="TemplateNotUnboundException"/>), not here — this method's only job is to attach and
    /// report, since it also runs for the "no template" and "clean template" paths where no such errors occur.
    /// </summary>
    private static (string StoreItemId, bool Replaced, ValidationResult Quick) AttachEverything(
        WordprocessingDocument doc, byte[] datasetBytes, ReportIdentity identity, string? headingText,
        bool scaffoldHeaderFooter)
    {
        var main = doc.MainDocumentPart!;
        var replaced = RemoveExistingBcParts(main);
        var storeItemId = AttachBcPart(main, datasetBytes, identity.Namespace);
        EnsureGlossaryPart(main);
        EnsureBodyHasContent(main, headingText ?? identity.ReportName);

        // BLANK builds only (see the parameter's own remarks): the empty header/footer parts every corpus
        // layout has, wired into the page setup EnsureBodyHasContent just established, so a from-scratch
        // layout can take a footer/header insert straight away instead of failing not_found with nothing to
        // resolve against. Done before the validator gate below, so a scaffolded part is
        // covered by the same structural check as everything else this method attaches.
        if (scaffoldHeaderFooter)
        {
            HeaderFooterScaffold.EnsureHeader(main);
            HeaderFooterScaffold.EnsureFooter(main);
        }

        main.Document!.Save();

        var openXmlErrors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc).ToList();
        if (openXmlErrors.Count > 0)
        {
            var preview = string.Join(" | ", openXmlErrors.Take(5).Select(e => $"{e.Path?.XPath}: {e.Description}"));
            throw new FileFormatException(
                $"The layout LayoutBuilder.Create just built failed OpenXmlValidator with {openXmlErrors.Count} "
                + "structural error(s); this indicates an internal bug (SdtFactory/LayoutBuilder are engineered "
                + $"to only ever emit valid OOXML), so nothing was written. First error(s): {preview}");
        }

        var quick = LayoutValidator.Quick(doc);
        return (storeItemId, replaced, quick);
    }

    // ---- schema source loading (.docx layout vs standalone schema .xml) ----

    /// <summary>
    /// Loads the raw dataset bytes plus report identity from <paramref name="schemaSource"/>: a <c>.docx</c>
    /// layout (bytes copied verbatim from its BC custom XML part; identity parsed via
    /// <see cref="SchemaProvider.FromLayout(WordprocessingDocument)"/> against the SAME open package) or a
    /// standalone schema <c>.xml</c> (bytes read as-is; identity — with a null <see cref="ReportIdentity.StoreItemId"/>,
    /// there being no OPC package/item props to read one from — parsed via <see cref="SchemaProvider.FromSchemaXml"/>,
    /// which also validates the root element).
    /// </summary>
    private static (byte[] DatasetBytes, ReportIdentity Identity) LoadDatasetSource(string schemaSource)
    {
        if (schemaSource.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
        {
            using var doc = WordprocessingDocument.Open(schemaSource, false);
            var main = doc.MainDocumentPart
                ?? throw new InvalidDataException($"'{schemaSource}' has no main document part.");

            var found = SchemaProvider.FindBcPart(main)
                ?? throw new InvalidDataException(
                    $"'{schemaSource}' has no BC dataset custom XML part (namespace starting "
                    + $"'{OoxmlNames.BcNamespacePrefix}').");

            // Parse identity from the SAME open package rather than reopening the file a second time.
            var tree = SchemaProvider.FromLayout(doc);

            using var partStream = found.Part.GetStream(FileMode.Open, FileAccess.Read);
            var bytes = ResourceLimits.ReadAllBytesCapped(
                partStream, $"Custom XML part '{PartWalker.PartFileName(found.Part)}' in '{schemaSource}'");
            return (bytes, tree.Report);
        }

        var schemaTree = SchemaProvider.FromSchemaXml(schemaSource);
        return (File.ReadAllBytes(schemaSource), schemaTree.Report);
    }

    // ---- BC custom XML part: remove-if-present, then attach verbatim bytes with a fresh itemID ----

    /// <summary>
    /// Removes every existing BC dataset custom XML part (and its own <see cref="CustomXmlPropertiesPart"/>)
    /// from <paramref name="main"/> — e.g. when starting from a <c>templatePath</c> that already has one.
    /// Any unrelated custom XML part (e.g. the Office bibliography part some corpus layouts carry) is left
    /// untouched, since <see cref="SchemaProvider.FindBcParts"/> filters by namespace, not by presence.
    /// Returns true when at least one BC part was found and removed.
    /// </summary>
    private static bool RemoveExistingBcParts(MainDocumentPart main)
    {
        var existing = SchemaProvider.FindBcParts(main);
        foreach (var part in existing)
        {
            // Delete the child properties part explicitly first rather than relying on DeletePart to
            // cascade — this is correct regardless of that behavior.
            if (part.CustomXmlPropertiesPart is { } props)
            {
                part.DeletePart(props);
            }

            main.DeletePart(part);
        }

        return existing.Count > 0;
    }

    /// <summary>
    /// Attaches <paramref name="datasetBytes"/> as a new BC dataset custom XML part — raw bytes, written
    /// verbatim to the part's stream (real BC parts are UTF-16 LE with a BOM; writing the exact source bytes
    /// rather than re-serializing preserves that exactly) — plus a <see cref="CustomXmlPropertiesPart"/>
    /// carrying a freshly generated <c>ds:itemID</c> GUID (<c>{GUID}</c>, uppercase — matching the corpus
    /// convention, e.g. <c>{AF7A6226-6056-400F-ADDA-E1ADA7C08250}</c>) and a <c>ds:schemaRef</c> pointing at
    /// <paramref name="ns"/>. Returns the new storeItemID.
    /// </summary>
    private static string AttachBcPart(MainDocumentPart main, byte[] datasetBytes, string ns)
    {
        var customXmlPart = main.AddCustomXmlPart(CustomXmlPartType.CustomXml);
        using (var partStream = customXmlPart.GetStream(FileMode.Create, FileAccess.Write))
        {
            partStream.Write(datasetBytes, 0, datasetBytes.Length);
        }

        var storeItemId = NewGuidToken();
        var propsPart = customXmlPart.AddNewPart<CustomXmlPropertiesPart>();
        propsPart.DataStoreItem = new Ds.DataStoreItem
        {
            ItemId = storeItemId,
            SchemaReferences = new Ds.SchemaReferences(new Ds.SchemaReference { Uri = ns }),
        };
        propsPart.DataStoreItem.Save();

        return storeItemId;
    }

    /// <summary><c>{GUID}</c>, uppercase — the corpus's own <c>ds:itemID</c>/docPart-guid convention.</summary>
    private static string NewGuidToken() => "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}";

    // ---- glossary part: ship the DefaultPlaceholder_-1854013440 docPart SdtFactory's placeholders reference ----

    /// <summary>
    /// Ensures <paramref name="main"/> has a <see cref="GlossaryDocumentPart"/> containing at least the
    /// built-in <c>DefaultPlaceholder_-1854013440</c> docPart entry every control <see cref="SdtFactory"/>
    /// builds references in its <c>w:placeholder</c> (see <see cref="SdtFactory.DefaultPlaceholderDocPart"/>),
    /// mirroring the real corpus glossary's own entry shape (element order: name, category(name, gallery),
    /// types(type), behaviors(behavior), guid) so Word/BC can resolve it. A no-op when a template's own
    /// glossary part already has that entry (true for every real corpus layout); when a template's glossary
    /// part exists but lacks it, the entry is appended without disturbing its other content.
    /// </summary>
    private static void EnsureGlossaryPart(MainDocumentPart main)
    {
        if (main.GlossaryDocumentPart is { } existing)
        {
            var hasDefaultPlaceholder = existing.GlossaryDocument?.Descendants<DocPartName>()
                .Any(n => n.Val?.Value == SdtFactory.DefaultPlaceholderDocPart) ?? false;
            if (hasDefaultPlaceholder)
            {
                return;
            }

            existing.GlossaryDocument ??= new GlossaryDocument();
            existing.GlossaryDocument.DocParts ??= new DocParts();
            existing.GlossaryDocument.DocParts.AppendChild(BuildDefaultPlaceholderDocPart());
            existing.GlossaryDocument.Save();
            return;
        }

        var glossaryPart = main.AddNewPart<GlossaryDocumentPart>();
        glossaryPart.GlossaryDocument = new GlossaryDocument(new DocParts(BuildDefaultPlaceholderDocPart()));
        glossaryPart.GlossaryDocument.Save();
    }

    /// <summary>
    /// Builds the <c>DefaultPlaceholder_-1854013440</c> docPart entry, mirroring
    /// <c>tests/corpus/SalesInvoiceForSubscriptionBilling.docx</c>'s <c>word/glossary/document.xml</c> element-for-element
    /// (only the entry's own <c>w:guid</c> differs, freshly generated here — Word's built-in entry's GUID has
    /// no meaning BC or this factory depends on; what matters is the <c>w:name</c>).
    /// </summary>
    private static DocPart BuildDefaultPlaceholderDocPart() => new(
        new DocPartProperties(
            new DocPartName { Val = SdtFactory.DefaultPlaceholderDocPart },
            new Category(new Name { Val = "General" }, new Gallery { Val = DocPartGalleryValues.Placeholder }),
            new DocPartTypes(new DocPartType { Val = DocPartValues.SdtPlaceholder }),
            new Behaviors(new Behavior { Val = DocPartBehaviorValues.Content }),
            new DocPartId { Val = NewGuidToken() }),
        new DocPartBody(new Paragraph(new Run(new Text("Click or tap here to enter text.")))));

    // ---- minimal body (fresh create) / heading-if-empty (template) ----

    /// <summary>
    /// Appends a bold heading paragraph carrying <paramref name="headingText"/> (the report name by
    /// default; <c>Create</c>'s optional <c>headingText</c> overrides it — an empty/whitespace override
    /// means NO heading, leaving one empty paragraph so the body stays well-formed) when
    /// <paramref name="main"/>'s body has no content of its own yet (a freshly created layout's blank
    /// <see cref="Body"/>, or a template whose body happened to already be empty) — otherwise leaves an
    /// existing (template) body exactly as it was: templates keep their own body. A trailing
    /// <see cref="SectionProperties"/> is added only when entirely absent; an existing one (with the
    /// template's own page size/margins/etc.) is always preserved rather than replaced.
    /// </summary>
    private static void EnsureBodyHasContent(MainDocumentPart main, string headingText)
    {
        var body = main.Document?.Body
            ?? throw new InvalidDataException("Layout has no document body.");

        var hasOtherContent = body.ChildElements.Any(e => e is not SectionProperties);
        if (!hasOtherContent)
        {
            var sectPr = body.Elements<SectionProperties>().FirstOrDefault();
            var heading = string.IsNullOrWhiteSpace(headingText)
                ? new Paragraph()
                : BuildHeadingParagraph(headingText);
            if (sectPr is not null)
            {
                body.InsertBefore(heading, sectPr);
            }
            else
            {
                body.AppendChild(heading);
            }
        }

        if (!body.Elements<SectionProperties>().Any())
        {
            // The BC-standard page setup every corpus layout shares: A4 (w:code 9) with 567-twip margins
            // and a 1134-twip left margin — an exact 10206-twip content width. Without this a blank
            // layout gets Word's defaults (Letter/1-inch margins, ~9360 twips), and every full-width
            // BC-style block authored into it overhangs the right margin. Only the blank path reaches
            // here; a template's own sectPr is always preserved above.
            body.AppendChild(new SectionProperties(
                new PageSize { Width = 11907, Height = 16839, Code = 9 },
                new PageMargin
                {
                    Top = 567,
                    Right = 567,
                    Bottom = 567,
                    Left = 1134,
                    Header = 567,
                    Footer = 567,
                    Gutter = 0,
                }));
        }
    }

    private static Paragraph BuildHeadingParagraph(string reportName) =>
        new(new Run(new RunProperties(new Bold(), new FontSize { Val = "32" }), new Text(reportName)));
}
