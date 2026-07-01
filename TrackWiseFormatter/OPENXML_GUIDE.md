# OpenXML Guide for the TrackWise Formatting Team

A focused, practical primer on the OpenXML SDK (`DocumentFormat.OpenXml`) for exactly the
operations this use case needs. Read this alongside
[`OpenXmlCookbook.cs`](src/Twf.DocEngine/Learning/OpenXmlCookbook.cs), which has a runnable
method for every technique below.

---

## 1. The mental model (this is what trips people up)

A `.docx` is **not** an XML file — it's a **ZIP package** (OPC) containing many XML **parts**
that reference each other through **relationships**. If you rename a `.docx` to `.zip` and
open it, you'll see:

```
/word/document.xml      <- the body text          (MainDocumentPart)
/word/header1.xml       <- a header               (HeaderPart)
/word/numbering.xml     <- list definitions       (NumberingDefinitionsPart)
/word/media/image1.png  <- an embedded image      (ImagePart)
/word/_rels/...         <- who references what    (relationships)
```

So in code you never "write XML text". You:
1. Open/create the **package** (`WordprocessingDocument`).
2. Add/get **parts** (`MainDocumentPart`, `HeaderPart`, `ImagePart`, ...).
3. Build a **strongly-typed element tree** inside each part.
4. Link parts with **relationship ids** (e.g. a header reference, an image blip).

## 2. The content tree

Inside `MainDocumentPart` the hierarchy is strict:

```
Document
└─ Body
   ├─ Paragraph (w:p)
   │  ├─ ParagraphProperties (w:pPr)   <- optional, MUST be first
   │  └─ Run (w:r)
   │     ├─ RunProperties (w:rPr)      <- optional, MUST be first
   │     └─ Text (w:t)                 <- set Space=Preserve to keep spaces
   ├─ Table (w:tbl)
   │  ├─ TableProperties (w:tblPr)
   │  └─ TableRow (w:tr) -> TableCell (w:tc) -> Paragraph (required!)
   └─ SectionProperties (w:sectPr)     <- MUST be the LAST child of Body
```

**Text lives only in `Text` inside `Run` inside `Paragraph`.** You cannot put text directly
in a cell or body — always wrap it in a paragraph → run → text.

## 3. The five ordering rules that cause "Word found unreadable content"

Almost every corruption bug is one of these:

| Rule | Why |
|---|---|
| `RunProperties` is the **first** child of a `Run` | schema order |
| `ParagraphProperties` is the **first** child of a `Paragraph` | schema order |
| Body-level `SectionProperties` (`sectPr`) is the **last** child of the body | it describes the section that ends there |
| `HeaderReference` comes **before** page size/margins inside `sectPr` | schema order |
| A `Table` declares its columns in a **`TableGrid`** right after `TableProperties` | a table with no `tblGrid` is invalid — Word repairs it |
| Every `TableCell` contains **at least one** `Paragraph`; a table is followed by a paragraph | Word requires it |

`OpenXmlCookbook` encodes all five so you don't have to memorise them — but now you know why
`AppendToBody()` inserts *before* the `sectPr`, and why `AddHeaderToAllPages()` uses
`PrependChild`.

## 4. The hard one: a header on every page (your #1 struggle)

This is three steps, not one — see `AddHeaderToAllPages()`:

```csharp
// (a) create the part, get its relationship id
var headerPart = main.AddNewPart<HeaderPart>();
var relId = main.GetIdOfPart(headerPart);

// (b) put content in it (must end with a paragraph)
headerPart.Header = new Header(identityTable, confidentialityParagraph, new Paragraph());

// (c) reference it from the section; Default type => every page
var sectPr = EnsureSectionProperties(body);
sectPr.PrependChild(new HeaderReference { Type = HeaderFooterValues.Default, Id = relId });
```

Gotchas:
- **Nothing shows** if you create the `HeaderPart` but forget step (c) — the reference is
  what binds it to the page.
- If a document has a **Title-page** or **even/odd** header, `Default` alone won't cover
  every page. For "same header everywhere", use only `Default` and don't set `TitlePage`.
- A header ending in a table (no trailing paragraph) renders as corrupt — always append a
  final empty paragraph.

## 5. Tables

`TableProperties` first, then rows. Borders are set once on the table via `TableBorders`
(six borders: top/bottom/left/right + insideH/insideV). Width `Type=Pct, Width="5000"` = 100%.
See `BorderedTable()`. Remember: **a cell needs a paragraph**, so we wrap every value in
`TextParagraph(...)`.

## 6. Images (the logo, or the "Exhibit A" flowchart)

Three steps again — see `AddInlineImage()`:
1. `AddImagePart(type)` creates `word/media/imageN.*`.
2. `imagePart.FeedData(stream)` streams the bytes in.
3. Wrap the relationship id in a `Drawing` → `Inline` → `Graphic` → `Picture` (the verbose
   DrawingML block). Size is in **EMU** (914,400 per inch; ~9,525 per pixel at 96 DPI).

## 7. Numbering (numeric section numbering)

List numbers are **not** literal text — they come from `numbering.xml`. You define an
`AbstractNum` (the format, e.g. `%1.` decimal) plus a `NumberingInstance` (a `numId`), then
attach the `numId` to each paragraph. See `DefineDecimalNumbering()` + `ApplyNumbering()`.
`AbstractNum` must appear before `Num` in the part. (For the Pilot we number the 8 sections
as plain bold text; real list numbering is here for when you tackle `I.A.1 → 1.1.1`.)

## 8. Two ways to build the output — and when to use each

| Approach | Class | Use when |
|---|---|---|
| **Populate the approved template** | `TrackWiseConverter` | Production — inherits the template's exact styles, logo, confidentiality header |
| **Build from scratch** | `Learning/ScratchDocumentWriter` | Learning + full control; no template file needed |

Both consume the same `SopDocument` model from `DocxParser`. Study `ScratchDocumentWriter`
to see the whole pipeline assembled with the cookbook; ship `TrackWiseConverter` for fidelity.

## 9. Open, edit, save — the lifecycle

```csharp
// create
using var doc = OpenXmlCookbook.CreateDocument("out.docx"); // written on dispose

// open existing for edit
using var doc = WordprocessingDocument.Open("in.docx", isEditable: true);
var body = doc.MainDocumentPart!.Document.Body!;
// ... mutate the tree ...
doc.MainDocumentPart.Document.Save();  // flush; file finalised on dispose
```

- The `using` block matters: the ZIP is only fully written when the document is disposed.
- `Open(path, false)` is read-only (parsing the source); `true` is editable.

## 10. How to learn it fastest (suggested order)

1. Run `ScratchDocumentWriter` once, open the result in Word — see a full document appear.
2. Rename that `.docx` to `.zip`, open `word/document.xml` — map what you see to the tree.
3. Read `OpenXmlCookbook` top to bottom; change one thing at a time and re-open.
4. Then read `DocxParser` (reading) and `TrackWiseConverter` (template-population).

## References

- OpenXML SDK docs: https://learn.microsoft.com/en-us/office/open-xml/word-processing
- The Open XML SDK "how do I..." snippets are the canonical source for the image/numbering blocks.
