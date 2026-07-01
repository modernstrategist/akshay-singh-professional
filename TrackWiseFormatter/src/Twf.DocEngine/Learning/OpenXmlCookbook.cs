using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace Twf.DocEngine.Learning;

/// <summary>
/// A teaching-oriented "cookbook" of the OpenXML techniques this use case needs.
/// Each method is small, self-contained, and heavily commented so the team can learn
/// the API one technique at a time. These are the exact operations people get stuck on.
///
/// MENTAL MODEL (read this first):
///   A .docx is a ZIP ("OPC package") of XML "parts" joined by "relationships".
///   WordprocessingDocument  -> the package
///     MainDocumentPart      -> word/document.xml   (the body text)
///       Document -> Body -> Paragraph -> Run -> Text     (the content tree)
///     HeaderPart            -> word/header1.xml    (referenced from the section)
///     ImagePart             -> word/media/image1.png (referenced by relationship id)
///     NumberingDefinitionsPart -> word/numbering.xml (list definitions)
///
/// ORDERING RULES that cause 80% of "corrupt document" errors:
///   * &lt;w:rPr&gt; (RunProperties) must be the FIRST child of a &lt;w:r&gt; (Run).
///   * &lt;w:pPr&gt; (ParagraphProperties) must be the FIRST child of a &lt;w:p&gt;.
///   * The body-level &lt;w:sectPr&gt; must be the LAST child of the body.
///   * &lt;w:headerReference&gt; must come BEFORE page size/margins inside &lt;w:sectPr&gt;.
///   * Every table cell must contain at least one paragraph; a table must be
///     followed by a paragraph.
/// </summary>
public static class OpenXmlCookbook
{
    // 914,400 English Metric Units (EMU) per inch — the unit DrawingML uses for image size.
    public const long EmuPerInch = 914_400L;

    // ---------------------------------------------------------------------
    // 1. CREATE a new document with an empty body.
    // ---------------------------------------------------------------------
    public static WordprocessingDocument CreateDocument(string path)
    {
        var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();          // creates word/document.xml
        main.Document = new Document(new Body());       // the root content tree
        return doc; // caller disposes -> that is when the file is actually written
    }

    // ---------------------------------------------------------------------
    // 2. A styled paragraph. Note RunProperties is appended FIRST inside the Run.
    // ---------------------------------------------------------------------
    public static Paragraph TextParagraph(
        string text, bool bold = false, JustificationValues? align = null, int? halfPointSize = null)
    {
        var run = new Run();

        var rpr = new RunProperties();
        if (bold) rpr.Append(new Bold());
        if (halfPointSize is int size) rpr.Append(new FontSize { Val = size.ToString() }); // half-points: 24 = 12pt
        if (rpr.HasChildren) run.Append(rpr);          // MUST be first child of the run

        run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve }); // Preserve keeps spaces

        var p = new Paragraph(run);
        if (align is JustificationValues j)
            p.ParagraphProperties = new ParagraphProperties(new Justification { Val = j }); // pPr first child
        return p;
    }

    // ---------------------------------------------------------------------
    // 3. A bordered table from rows of strings. First row can be bold (header).
    // ---------------------------------------------------------------------
    public static Table BorderedTable(IEnumerable<string[]> rows, bool firstRowBold = true)
    {
        var table = new Table(new TableProperties(
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4U },
                new BottomBorder { Val = BorderValues.Single, Size = 4U },
                new LeftBorder { Val = BorderValues.Single, Size = 4U },
                new RightBorder { Val = BorderValues.Single, Size = 4U },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4U },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4U }),
            new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" })); // 5000 = 100%

        bool isFirst = true;
        foreach (var row in rows)
        {
            var tr = new TableRow();
            foreach (var cellText in row)
            {
                // A cell MUST contain at least one paragraph, or Word reports corruption.
                tr.Append(new TableCell(TextParagraph(cellText, bold: firstRowBold && isFirst)));
            }
            table.Append(tr);
            isFirst = false;
        }
        return table;
    }

    // ---------------------------------------------------------------------
    // 4. THE HARD ONE: a header that appears on EVERY page, built from scratch.
    //    Creates the HeaderPart, then references it from the section properties.
    // ---------------------------------------------------------------------
    public static string AddHeaderToAllPages(WordprocessingDocument doc, params OpenXmlElement[] content)
    {
        var main = doc.MainDocumentPart!;

        // (a) create the part and get its relationship id
        var headerPart = main.AddNewPart<HeaderPart>();
        var relId = main.GetIdOfPart(headerPart);

        // (b) fill it; a header must end with a paragraph, so append a trailing one
        var header = new Header();
        foreach (var el in content) header.Append(el);
        header.Append(new Paragraph());
        headerPart.Header = header;

        // (c) reference it from the section. Type=Default => shows on every page
        //     (as long as there is no separate Title/Even-page header).
        var sectPr = EnsureSectionProperties(main.Document.Body!);
        sectPr.RemoveAllChildren<HeaderReference>();
        sectPr.PrependChild(new HeaderReference          // MUST be first inside sectPr
        {
            Type = HeaderFooterValues.Default,
            Id = relId
        });
        return relId;
    }

    // ---------------------------------------------------------------------
    // 5. Section properties: create if missing, always kept LAST in the body.
    // ---------------------------------------------------------------------
    public static SectionProperties EnsureSectionProperties(Body body)
    {
        var sectPr = body.Elements<SectionProperties>().LastOrDefault();
        if (sectPr is null)
        {
            sectPr = new SectionProperties();
            body.Append(sectPr); // last child
        }
        return sectPr;
    }

    /// <summary>Append body content while guaranteeing the sectPr stays last.</summary>
    public static void AppendToBody(Body body, OpenXmlElement element)
    {
        var sectPr = body.Elements<SectionProperties>().LastOrDefault();
        if (sectPr is not null) body.InsertBefore(element, sectPr);
        else body.Append(element);
    }

    // ---------------------------------------------------------------------
    // 6. Inline image (e.g. the logo, or the legacy "Exhibit A" flowchart).
    //    Teaches ImagePart + relationship + the DrawingML wrapper.
    //    cx/cy are in EMU: use EmuPerInch, or ~9525 EMU per pixel at 96 DPI.
    // ---------------------------------------------------------------------
    public static void AddInlineImage(
        WordprocessingDocument doc, Paragraph target, Stream imageStream,
        ImagePartType type, long cx, long cy, string name = "image")
    {
        var main = doc.MainDocumentPart!;
        var imagePart = main.AddImagePart(type);   // creates word/media/imageN.*
        imagePart.FeedData(imageStream);            // stream the bytes into the part
        var relId = main.GetIdOfPart(imagePart);    // relationship id used by the Blip

        target.Append(new Run(BuildInlineDrawing(relId, cx, cy, name)));
    }

    // The canonical DrawingML "inline picture" wrapper. Verbose but mechanical.
    private static Drawing BuildInlineDrawing(string relId, long cx, long cy, string name) =>
        new(new DW.Inline(
                new DW.Extent { Cx = cx, Cy = cy },
                new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                new DW.DocProperties { Id = 1U, Name = name },
                new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks { NoChangeAspect = true }),
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties { Id = 0U, Name = name },
                                new PIC.NonVisualPictureDrawingProperties()),
                            new PIC.BlipFill(
                                new A.Blip { Embed = relId, CompressionState = A.BlipCompressionValues.Print },
                                new A.Stretch(new A.FillRectangle())),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = 0L, Y = 0L },
                                    new A.Extents { Cx = cx, Cy = cy }),
                                new A.PresetGeometry(new A.AdjustValueList())
                                { Preset = A.ShapeTypeValues.Rectangle })))
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
            { DistanceFromTop = 0U, DistanceFromBottom = 0U, DistanceFromLeft = 0U, DistanceFromRight = 0U });

    // ---------------------------------------------------------------------
    // 7. Decimal list numbering (relevant to numeric section numbering).
    //    Returns a numId you attach to paragraphs via ApplyNumbering().
    // ---------------------------------------------------------------------
    public static int DefineDecimalNumbering(WordprocessingDocument doc, int abstractNumId = 1, int numId = 1)
    {
        var main = doc.MainDocumentPart!;
        var part = main.NumberingDefinitionsPart ?? main.AddNewPart<NumberingDefinitionsPart>();
        part.Numbering ??= new Numbering();

        // Level 0: "1." "2." "3." ...
        var abstractNum = new AbstractNum(
            new Level(
                new StartNumberingValue { Val = 1 },
                new NumberingFormat { Val = NumberFormatValues.Decimal },
                new LevelText { Val = "%1." },
                new LevelJustification { Val = LevelJustificationValues.Left })
            { LevelIndex = 0 })
        { AbstractNumberId = abstractNumId };

        // AbstractNum elements must appear BEFORE Num elements in numbering.xml.
        part.Numbering.InsertAt(abstractNum, 0);
        part.Numbering.Append(new NumberingInstance(new AbstractNumId { Val = abstractNumId }) { NumberID = numId });
        return numId;
    }

    public static void ApplyNumbering(Paragraph paragraph, int numId, int level = 0)
    {
        paragraph.ParagraphProperties ??= new ParagraphProperties();
        paragraph.ParagraphProperties.NumberingProperties =
            new NumberingProperties(new NumberingLevelReference { Val = level }, new NumberingId { Val = numId });
    }
}
