# TrackWise Formatting — Pilot (Phase 1)

A .NET 8 class-library solution that converts a legacy IVC SOP (Word `.docx`) into the
new **TrackWise (TWD) SOP format**, using **OpenXML** for deterministic document
generation. This is the Phase 1 / "can it be done?" feasibility build, proven end-to-end
on the real **GEN-027 — Handling of Manufacturing Reworks** DDR sample.

> Design principle: **deterministic code moves all regulated SOP content.** No LLM
> generates document body text. AI (Azure AI Foundry / OpenAI) is reserved for the
> fuzzy section-mapping assist added in Phase 3.

## Learning OpenXML (start here if the SDK is new to the team)

- **[`OPENXML_GUIDE.md`](OPENXML_GUIDE.md)** — a focused primer: the package/part/relationship
  mental model, the content tree, the five ordering rules that cause "unreadable content",
  and step-by-step for headers/tables/images/numbering.
- **[`src/Twf.DocEngine/Learning/OpenXmlCookbook.cs`](src/Twf.DocEngine/Learning/OpenXmlCookbook.cs)**
  — one small, heavily-commented method per technique (create doc, styled paragraph, bordered
  table, **header on every page**, section properties, inline image, list numbering).
- **[`src/Twf.DocEngine/Learning/ScratchDocumentWriter.cs`](src/Twf.DocEngine/Learning/ScratchDocumentWriter.cs)**
  — builds the whole TrackWise document **from scratch** (no template needed) so you can see
  the full pipeline assembled. Contrast with `TrackWiseConverter`, which populates the
  approved template for production fidelity.

- **[`OPENXML_CHEATSHEET.md`](OPENXML_CHEATSHEET.md)** — a one-page printable reference of the
  most-used snippets.

Instant demo (no client files needed):
```bash
dotnet run --project src/Twf.Console -- --scratch demo.docx
```
This builds a full TrackWise SOP from scratch so the team gets an open-in-Word artifact while learning.

Suggested path: run the `--scratch` demo → open the result in Word → rename it to `.zip`
and read `word/document.xml` → then read the cookbook top to bottom.

## What it does

Given a legacy SOP + the approved `NEW TWD SOP template.docx`, it produces a **draft**
output document that:

1. Rebuilds the **header** — Department, **Document Number** (looked up from the legacy
   ID, e.g. `GEN-027 → SOP-DDR-GEN-000003`), incremented **Revision Number**, Title — with
   the **confidentiality statement on every page** (from the template header).
2. Emits all **eight mandatory sections** (Purpose, Scope, Responsibilities, References,
   Definitions, Procedure, Forms, Revision History), numbered `1.`–`8.`, with **N/A** for
   any section absent from the source.
3. Maps legacy content into the target sections (e.g. `RESPONSIBILITY → Responsibilities`;
   merges split `PROCEDURE` / `PROCEDURE: (continued)` blocks).
4. Rebuilds **Revision History** as a **4-line rolling table** — keeps the last 3 prior
   revisions, increments the number, adds a new row with `(see header)` and reason
   *"Updated to TrackWise format."*
5. Produces an **Exception Report** for anything deferred or unresolved.
6. **Never modifies the source** — it builds against a copy of the template.

## Project structure

| Project | Responsibility |
|---|---|
| `Twf.Core` | Domain models, `ConversionOptions`, `ConversionResult` — no dependencies |
| `Twf.DocEngine` | The OpenXML engine: `DocxParser`, `HeaderBuilder`, `SectionBuilder`, `RevisionHistoryBuilder`, `LegacyIdMap`, `TrackWiseConverter` |
| `Twf.Console` | CLI runner that converts one document and prints the exception report |

## Build & run

```bash
dotnet build TrackWiseFormatter.sln

dotnet run --project src/Twf.Console -- \
    "SOPDDRGEN000003_GEN027_v5_Handling_of_Manufacturing_Reworks.docx" \
    "NEW_TWD_SOP_template.docx" \
    "out/SOP-DDR-GEN-000003_DRAFT.docx" \
    samples/legacy-ids.csv
```

The only NuGet dependency is `DocumentFormat.OpenXml` (3.1.0).

## Library usage

```csharp
var legacy = LegacyIdMap.FromCsv("samples/legacy-ids.csv");
var converter = new TrackWiseConverter(legacy, new ConversionOptions { RevisionPadding = 3 });
var result = converter.Convert(sourcePath, templatePath, outputPath, DocType.Ddr);

foreach (var ex in result.Exceptions)
    Console.WriteLine($"[{ex.Severity}] {ex.Message}");
```

## Open decisions (pin these in the Phase 1 spec / D1.4)

These come straight from comparing the real GEN-027 source with the template. They are
exposed as `ConversionOptions` so the decision is config, not a code change:

| Decision | Default here | Conflict |
|---|---|---|
| Revision-number padding | `3` → `006` | SOW §5.2 says "omit leading zeroes" (`6`) |
| Title casing | Title Case | Source is ALL CAPS |
| Exhibits / flowchart images | Deferred to Phase 3 | Not specified in SOW |

## Mapping to the Pilot (100-hour) task list

| Task | Where it lives |
|---|---|
| OpenXML ramp-up / spike | `OpenXmlHelpers`, this engine |
| Minimal .NET scaffold | the `.sln` + 3 projects |
| `DocxParser` | `Twf.DocEngine/DocxParser.cs` |
| `HeaderBuilder` | `Twf.DocEngine/HeaderBuilder.cs` |
| `SectionBuilder` | `Twf.DocEngine/SectionBuilder.cs` |
| `RevisionHistoryBuilder` | `Twf.DocEngine/RevisionHistoryBuilder.cs` |
| OutputWriter + exception list | `TrackWiseConverter` + `Twf.Console` |
| Demo + findings | `FINDINGS.md`, sample output in `samples/` |

## Deferred (Phase 2 / 3 / Production)

- NTR document type (strip Effective Date), legacy-ID lookup at scale via Azure SQL
- Procedure sub-step numbering normalization (`I.A.1 → 1.1.1`)
- Image / Exhibit carry-over
- AI-assisted mapping for non-standard headings (Azure AI Foundry)
- App Service UI, Azure Functions orchestration, Blob in/out directories
- Bulk conversion, TrackWise integration, GxP/Part 11 validation
