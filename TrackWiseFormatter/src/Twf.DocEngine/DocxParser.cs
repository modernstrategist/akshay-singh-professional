using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Twf.Core;

namespace Twf.DocEngine;

/// <summary>
/// Reads a legacy SOP .docx into the internal <see cref="SopDocument"/> model:
/// header identity fields, the eight target sections (by heading alias), and the
/// legacy revision-history table. Purely deterministic — no LLM involved.
/// </summary>
public sealed class DocxParser
{
    // Heading recognition is shared with TemplateFiller via SectionHeadings.

    public SopDocument Parse(string sourcePath, DocType docType = DocType.Ddr)
    {
        using var src = WordprocessingDocument.Open(sourcePath, isEditable: false);
        var body = src.MainDocumentPart!.Document.Body!;
        var model = new SopDocument { DocType = docType };

        // Every target section exists up front; empty ones become N/A downstream.
        foreach (TargetSection ts in Enum.GetValues<TargetSection>())
            model.Sections[ts] = new SopSection { Target = ts };

        ParseHeader(src, model);
        ParseBody(body, model);
        ParseRevisionTable(body, model);
        return model;
    }

    private static void ParseHeader(WordprocessingDocument src, SopDocument model)
    {
        foreach (var hp in src.MainDocumentPart!.HeaderParts)
        {
            foreach (var raw in OpenXmlHelpers.VisibleLines(hp.Header))
            {
                var line = raw.Trim();
                if (line.Length == 0) continue;

                if (line.StartsWith("SOP Number", StringComparison.OrdinalIgnoreCase))
                    model.LegacyId = AfterColon(line);
                else if (line.StartsWith("Subject", StringComparison.OrdinalIgnoreCase))
                    model.Title = AfterColon(line);

                var m = Regex.Match(line, @"revision\s+number\s*:?\s*(\d+)", RegexOptions.IgnoreCase);
                if (m.Success) model.OldRevisionNumber = int.Parse(m.Groups[1].Value);
            }
        }
    }

    private static void ParseBody(Body body, SopDocument model)
    {
        TargetSection? current = null;
        foreach (var p in body.Elements<Paragraph>())
        {
            var text = p.InnerText.Trim();
            if (text.Length == 0) continue;

            // A short line that matches a known heading switches the current section.
            if (text.Length < 40 && SectionHeadings.TryMatch(text, out var tgt))
            {
                // Revision History content comes from the table, not body paragraphs.
                current = tgt == TargetSection.RevisionHistory ? null : tgt;
                continue;
            }

            if (current is { } c)
                model.Sections[c].Paragraphs.Add(p.InnerText);
        }
    }

    private static void ParseRevisionTable(Body body, SopDocument model)
    {
        var table = body.Elements<Table>().FirstOrDefault();
        if (table is null) return;

        foreach (var row in table.Elements<TableRow>().Skip(1)) // skip header row
        {
            var cells = row.Elements<TableCell>().Select(c => c.InnerText.Trim()).ToList();
            if (cells.Count >= 3 && cells.Any(x => x.Length > 0))
                model.RevisionHistory.Add(new RevisionEntry
                {
                    RevisionNumber = cells[0],
                    EffectiveDate = cells[1],
                    Reason = cells[2]
                });
        }
    }

    private static string AfterColon(string line)
    {
        var i = line.IndexOf(':');
        return i >= 0 ? line[(i + 1)..].Trim() : line.Trim();
    }
}
