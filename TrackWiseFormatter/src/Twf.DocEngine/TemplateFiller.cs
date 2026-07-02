using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Twf.Core;
using static Twf.DocEngine.OpenXmlHelpers;

namespace Twf.DocEngine;

/// <summary>
/// The "work on the OUTCOME, not the input" strategy.
///
/// The approved output template already defines the requested FORMAT — the eight headings
/// (exactly as they should read), the FORMS layout, and a valid Revision History table.
/// This class keeps all of that and simply INSERTS the mapped source content under each
/// existing heading, and fills the template's own revision table. Because we never rebuild
/// the headings or the table, the output is guaranteed to match the requested format
/// (no re-numbered headings, no missing &lt;w:tblGrid&gt;).
/// </summary>
public sealed class TemplateFiller
{
    public void Fill(WordprocessingDocument doc, SopDocument model, ConversionOptions options,
        ReviewAnnotator? annotator, System.DateTime stampUtc)
    {
        var body = doc.MainDocumentPart!.Document.Body!;
        FillSections(doc, body, model, options, annotator, stampUtc);
        FillRevisionTable(body, model, options);
    }

    private static void FillSections(WordprocessingDocument doc, Body body, SopDocument model,
        ConversionOptions options, ReviewAnnotator? annotator, System.DateTime stampUtc)
    {
        // Snapshot of block-level elements in document order.
        var blocks = body.ChildElements.Where(e => e is Paragraph or Table).ToList();

        // Locate each heading paragraph the template already provides.
        var headings = new List<(TargetSection Target, Paragraph Heading, int Index)>();
        for (int i = 0; i < blocks.Count; i++)
        {
            if (blocks[i] is Paragraph p)
            {
                var text = p.InnerText.Trim();
                if (text.Length is > 0 and < 60 && SectionHeadings.TryMatch(text, out var tgt))
                    headings.Add((tgt, p, i));
            }
        }

        foreach (var (target, heading, index) in headings)
        {
            if (target == TargetSection.RevisionHistory) continue; // table handled separately

            // Region for this section = everything up to the next heading (or end).
            int boundary = headings.Where(h => h.Index > index)
                                   .Select(h => h.Index)
                                   .DefaultIfEmpty(blocks.Count)
                                   .Min();

            // Remove the template's placeholder paragraphs in this region (keep any table).
            for (int j = index + 1; j < boundary; j++)
                if (blocks[j] is Paragraph placeholder)
                    placeholder.Remove();

            // Insert the mapped content immediately under the heading (or N/A if none).
            OpenXmlElement anchor = heading;
            var section = model.Sections.TryGetValue(target, out var s) ? s : null;
            if (section is null || section.IsEmpty)
            {
                var na = Para("N/A", highlightYellow: options.EnableHighlighting);
                body.InsertAfter(na, anchor);
                annotator?.AddComment(doc, na,
                    $"{target}: section not present in source; set to N/A for review.", stampUtc);
            }
            else
            {
                foreach (var line in section.Paragraphs)
                    anchor = body.InsertAfter(Para(line), anchor);
            }
        }
    }

    // Reuse the template's existing revision table so its tblPr + tblGrid are preserved.
    private static void FillRevisionTable(Body body, SopDocument model, ConversionOptions options)
    {
        var table = body.Elements<Table>().FirstOrDefault();
        if (table is null) return;

        var rows = table.Elements<TableRow>().ToList();
        var headerRow = rows.FirstOrDefault();       // keep column header row
        foreach (var r in rows.Skip(1)) r.Remove();  // clear placeholder data rows

        var prior = model.RevisionHistory.TakeLast(Math.Max(0, options.PriorRevisionsToKeep));
        var newRev = HeaderBuilder.FormatRevision(model.NewRevisionNumber, options.RevisionPadding);

        foreach (var e in prior)
            table.Append(BuildRow(headerRow, e.RevisionNumber, e.EffectiveDate, e.Reason));
        table.Append(BuildRow(headerRow, newRev, "(see header)", options.NewRevisionReason));
    }

    // Clone the header row's cell properties so column widths are preserved.
    private static TableRow BuildRow(TableRow? templateRow, params string[] values)
    {
        var templateCells = templateRow?.Elements<TableCell>().ToList();
        var tr = new TableRow();
        for (int i = 0; i < values.Length; i++)
        {
            var cell = new TableCell();
            if (templateCells is not null && i < templateCells.Count &&
                templateCells[i].GetFirstChild<TableCellProperties>()?.CloneNode(true) is TableCellProperties tcPr)
                cell.Append(tcPr);
            cell.Append(Para(values[i]));
            tr.Append(cell);
        }
        return tr;
    }
}
