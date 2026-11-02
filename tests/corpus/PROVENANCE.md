# Corpus provenance

This directory holds real Business Central Word-layout captures. They are the project's primary
fixtures precisely because they are genuine BC output — warts included — so they are never
regenerated, "cleaned up", or re-saved through Word. The two records below exist so the questions
they answer do not get re-litigated later.

## What these files are, and why they may be published

Each file is a Word report layout as exported from / shipped with Business Central. Their static
content was reviewed before publication (2026-08-01): every document body consists of
data-binding placeholders and caption names only — no Fusion5 data, no client data, no personal
data, no document text beyond the layout's own captions.

The OOXML **metadata** of the original captures did carry personal residue (author/last-saved-by
names, MSIP sensitivity-label parts with tenant GUIDs and one personal e-mail address, external
`attachedTemplate` relationships pointing into named developers' profile folders, and in one file a
flat-OPC clipboard accident embedded in a dataset sample value). All of it was removed on
2026-08-01 by [`tools/scrub_corpus_metadata.py`](../../tools/scrub_corpus_metadata.py) — see its
docstring for the exact rules, including the two deliberate softenings: external `attachedTemplate`
relationships are kept with a sanitized `C:\Templates\` target (the validator's `attached-template` check,
the tools' unlinking of them and the preview-strip tests exercise them), and everything else in the packages — BC dataset parts,
bodies, media, styles — is byte-identical to the capture.

The three themable layouts (`*Themable.docx`) joined on 2026-10-29. They are Microsoft's own stock
layouts, so their content is placeholders and captions only, and the same script scrubbed them before
they were committed: author/last-saved-by names, the MSIP label parts in
`StandardPurchaseOrderThemable.docx`, and `attachedTemplate` targets in developers' Desktop folders.

`ExternalDefault.docx` and `SpendRequestDocument.docx` joined on 2026-11-03 as the first tables that merge
cells vertically (`w:vMerge`, GitHub issue #9). They are Microsoft's own layouts too, placeholders and
captions only, scrubbed the same way: author/last-saved-by names, the MSIP label parts, and in
`ExternalDefault.docx` an empty bibliography part.

**Run the script whenever a new capture joins the corpus, before it is committed.**

| File | Report (dataset namespace) | Origin |
|---|---|---|
| `ExternalDefault.docx` | `Tenant_Report_Defaults/2000000001` (header/footer design *External Default*) | Microsoft base app |
| `InventoryOrderDetails.docx` | `FS_YSR_InventoryOrderDetails/50002` | Custom-report capture (content: placeholders/captions only) |
| `JobQuote.docx` | `Job_Quote/1016` | Microsoft base app |
| `PaymentPracticeByPeriod.docx` | `Payment_Practice/685` | Microsoft Payment Practices app |
| `SalesInvoiceForSubscriptionBilling.docx` | `Standard_Sales_Invoice/1306` | Microsoft Subscription Billing app (report-extension layout) |
| `SalespersonCommission.docx` | `Salesperson_Commission/115` | Microsoft base app |
| `SpendRequestDocument.docx` | `Spend_Request_Document/6850` | Microsoft base app |
| `StandardPurchaseOrder.docx` | `Standard_Purchase_Order/1322` | Microsoft base app |
| `StandardPurchaseOrderThemable.docx` | `Standard_Purchase_Order/1322` (themable variant) | Microsoft base app |
| `StandardSalesInvoiceBlueSimpleThemable.docx` | `Standard_Sales_Invoice/1306` (Blue Simple, themable variant) | Microsoft base app |
| `StandardSalesInvoiceVatSpec.docx` | `Standard_Sales_Invoice/1306` (VAT-spec variant) | Microsoft base app |
| `StandardSalesQuote.docx` | `Standard_Sales_Quote/1304` | Microsoft base app |
| `StandardSalesShipmentBlueThemable.docx` | `Standard_Sales_Shipment/1308` (Blue, themable variant) | Microsoft base app |
| `StandardStatement.docx` | `Standard_Statement/1316` (no embedded dataset part) | Microsoft base app |
| `SubcontractorDispatchList.docx` | `Subcontractor_Dispatch_List/99000789` | Microsoft base app |

(Namespace prefixes and object-id ranges are **not** provenance evidence — Microsoft base-app
layouts themselves carry `FS_`/50000-range namespaces in places. The origin column above records
what each file actually is, established during the pre-publication review.)

## Position on redistributing the Microsoft-origin layouts

Recorded 2026-08-01. The nine Microsoft-origin files of that date are captures of report layouts whose
source files Microsoft publishes as open source in
[`microsoft/BCApps`](https://github.com/microsoft/BCApps) under the **MIT License**
(Copyright (c) Microsoft Corporation). Verified against that repository on the same date, per file
(the three `*Themable.docx` rows were added on 2026-10-29, and `ExternalDefault.docx` and
`SpendRequestDocument.docx` on 2026-11-03, see below):

| Corpus file | Source in `microsoft/BCApps` |
|---|---|
| `ExternalDefault.docx` | `src/Layers/W1/BaseApp/Foundation/Reporting/ReportParts/HeaderFooterDesign/ExternalDefault.docx` |
| `JobQuote.docx` | `src/Layers/W1/BaseApp/Projects/Project/JobQuote.docx` |
| `PaymentPracticeByPeriod.docx` | `src/Apps/W1/PaymentPractices/App/src/Reports/Payment Practice by Period.docx` |
| `SalesInvoiceForSubscriptionBilling.docx` | `src/Apps/W1/Subscription Billing/App/Billing/Report Extensions/Layouts/SalesInvoiceForSubscriptionBilling.docx` |
| `SalespersonCommission.docx` | `src/Layers/W1/BaseApp/Sales/Reports/SalespersonCommission.docx` |
| `SpendRequestDocument.docx` | `src/Layers/W1/BaseApp/Finance/SpendRequest/SpendRequestDocument.docx` |
| `StandardPurchaseOrder.docx` | `src/Layers/W1/BaseApp/Purchases/Document/StandardPurchaseOrder.docx` |
| `StandardPurchaseOrderThemable.docx` | `src/Layers/W1/BaseApp/Purchases/Document/StandardPurchaseOrderThemable.docx` |
| `StandardSalesInvoiceBlueSimpleThemable.docx` | `src/Layers/W1/BaseApp/Sales/History/StandardSalesInvoiceBlueSimpleThemable.docx` |
| `StandardSalesInvoiceVatSpec.docx` | `src/Layers/W1/BaseApp/Sales/History/StandardSalesInvoiceVatSpec.docx` |
| `StandardSalesQuote.docx` | `src/Layers/W1/BaseApp/Sales/Document/StandardSalesQuote.docx` |
| `StandardSalesShipmentBlueThemable.docx` | `src/Layers/W1/BaseApp/Sales/History/StandardSalesShipmentBlueThemable.docx` |
| `StandardStatement.docx` | `src/Layers/W1/BaseApp/Sales/Customer/StandardStatement.docx` |
| `SubcontractorDispatchList.docx` | `src/Layers/W1/BaseApp/Manufacturing/Reports/SubcontractorDispatchList.docx` |

The three themable files are the source files themselves, not BC-served captures. Before the scrub
they were byte-identical to the files above at BCApps `main` @ `0a4207a440` (git blob hashes compared
on 2026-10-29). Microsoft publishes further themable layouts beside them; those were left out because they
repeat what these three already show.

`ExternalDefault.docx` and `SpendRequestDocument.docx` are source files too: before the scrub they were
byte-identical to the files above at BCApps `main` @ `19ed5d01ff` (git blobs `d9edfa0d68` and `064c324402`,
2026-11-03). Thirteen of the 368 layouts in BCApps merge cells vertically: nine BC 29 header/footer designs,
`SpendRequestDocument.docx` and three NA demo invoices. `ExternalDefault.docx` stands for the designs, whose
header tables (those checked) merge the report-caption column down every row; `SpendRequestDocument.docx` merges
a `gridSpan` 2 cell over two rows of a `w:gridAfter` grid. The demo invoices merge both side columns of their
address block down nine rows, which is the first shape twice.

Redistribution here is under those MIT terms: Microsoft's copyright and licence notice for these
files is carried in `THIRD-PARTY-NOTICES.md`. What this repository redistributes is either a
BC-served capture (the serving environment re-saves parts of the package) or, for the themable
files, the source file; the metadata scrub above applies to both. The MIT licence permits this
(modification and redistribution, with the notice retained).

`InventoryOrderDetails.docx` is not from BCApps: it is a capture of a custom report layout, our own
material (cleared by the same pre-publication review), published under this repository's own MIT licence.

## Local-only material

Anything under `tests/corpus/protected-private/` is private reference material: gitignored, never
tracked (verified against the full history on 2026-08-01), never published, and never referenced
from tracked code or docs.
