# OpenXML Cheat-Sheet (one page)

`using DocumentFormat.OpenXml; using DocumentFormat.OpenXml.Packaging; using DocumentFormat.OpenXml.Wordprocessing;`

### Open / create / save
```csharp
using var doc = WordprocessingDocument.Create("out.docx", WordprocessingDocumentType.Document);
var main = doc.AddMainDocumentPart();
main.Document = new Document(new Body());
// ...build...
main.Document.Save();                 // file is finalised when `doc` is disposed

using var doc = WordprocessingDocument.Open("in.docx", true);   // true = editable
var body = doc.MainDocumentPart!.Document.Body!;
```

### Paragraph + styled run  (rPr MUST be first child of the run)
```csharp
var run = new Run(new RunProperties(new Bold(), new FontSize { Val = "24" }),   // 24 = 12pt
                  new Text("Hello") { Space = SpaceProcessingModeValues.Preserve });
var p = new Paragraph(run);
```

### Add to body but keep sectPr last
```csharp
var sectPr = body.Elements<SectionProperties>().LastOrDefault();
if (sectPr != null) body.InsertBefore(p, sectPr); else body.Append(p);
```

### Bordered table (cell MUST contain a paragraph)
```csharp
var t = new Table(new TableProperties(
    new TableBorders(
        new TopBorder{Val=BorderValues.Single,Size=4U}, new BottomBorder{Val=BorderValues.Single,Size=4U},
        new LeftBorder{Val=BorderValues.Single,Size=4U}, new RightBorder{Val=BorderValues.Single,Size=4U},
        new InsideHorizontalBorder{Val=BorderValues.Single,Size=4U}, new InsideVerticalBorder{Val=BorderValues.Single,Size=4U}),
    new TableWidth{Type=TableWidthUnitValues.Pct, Width="5000"}));      // 5000 = 100%
t.Append(new TableRow(new TableCell(new Paragraph(new Run(new Text("cell"))))));
```

### Header on EVERY page (3 steps — this is the one people miss)
```csharp
var hp   = main.AddNewPart<HeaderPart>();               // 1) create part
var id   = main.GetIdOfPart(hp);
hp.Header = new Header(myTable, myParagraph, new Paragraph());  // 2) content (end with a paragraph)

var sectPr = body.Elements<SectionProperties>().LastOrDefault() ?? body.AppendChild(new SectionProperties());
sectPr.PrependChild(new HeaderReference { Type = HeaderFooterValues.Default, Id = id }); // 3) bind (ref first in sectPr)
```

### Inline image  (EMU: 914400 per inch, ~9525 per px @96dpi)
```csharp
var img = main.AddImagePart(ImagePartType.Png);
using (var fs = File.OpenRead("logo.png")) img.FeedData(fs);
var relId = main.GetIdOfPart(img);
// then wrap relId in Drawing>Inline>Graphic>Picture (see OpenXmlCookbook.AddInlineImage)
```

### Read text from an existing doc
```csharp
foreach (var p in body.Elements<Paragraph>()) Console.WriteLine(p.InnerText);
foreach (var tbl in body.Elements<Table>())
  foreach (var row in tbl.Elements<TableRow>())
    Console.WriteLine(string.Join(" | ", row.Elements<TableCell>().Select(c => c.InnerText)));
```

### The 5 rules that prevent "unreadable content"
1. `RunProperties` first inside `Run`. 2. `ParagraphProperties` first inside `Paragraph`.
3. body `SectionProperties` is **last**. 4. `HeaderReference` first inside `sectPr`.
5. Every cell has a paragraph; a table is followed by a paragraph.

### Debug tip
Rename any `.docx` to `.zip`, open `word/document.xml` — the XML you see maps 1:1 to the element tree you build in code.

---
Full explanations: **OPENXML_GUIDE.md** · Runnable methods: **src/Twf.DocEngine/Learning/OpenXmlCookbook.cs**
Try it now: `dotnet run --project src/Twf.Console -- --scratch demo.docx`
