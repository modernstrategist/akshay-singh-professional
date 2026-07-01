using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Twf.Core;
using static Twf.DocEngine.Learning.OpenXmlCookbook;

namespace Twf.DocEngine.Learning;

/// <summary>
/// Builds the new TrackWise SOP document ENTIRELY FROM SCRATCH — no template file needed.
/// This is the teaching counterpart to <see cref="TrackWiseConverter"/> (which populates a
/// copy of the approved template). Use this to learn how every part is assembled in OpenXML;
/// use the template-based converter in production when you need pixel-exact fidelity to the
/// approved template's styles.
///
/// Pipeline demonstrated here:
///   CreateDocument -> build+attach the every-page header -> write 8 sections -> revision table -> save.
/// </summary>
public static class ScratchDocumentWriter
{
    private const string ConfidentialityStatement =
        "Confidentiality Statement:  Standard Operating Procedure (SOP) documents are confidential. " +
        "These documents are controlled and are not to be copied or made available to personnel " +
        "without notifying IVC Management.";

    public static void Write(SopDocument model, ConversionOptions options, string outputPath)
    {
        using var doc = CreateDocument(outputPath);
        var main = doc.MainDocumentPart!;
        var body = main.Document.Body!;

        BuildHeader(doc, model, options);
        BuildBody(body, model, options);

        main.Document.Save(); // flush the content tree (file is finalised on dispose)
    }

    // Header = identity table + confidentiality statement, shown on every page.
    private static void BuildHeader(WordprocessingDocument doc, SopDocument model, ConversionOptions options)
    {
        var rev = HeaderBuilder.FormatRevision(model.NewRevisionNumber, options.RevisionPadding);

        var identity = BorderedTable(new[]
        {
            new[] { "Standard Operating Procedure", "", "" },
            new[] { $"Department:  {model.Department}", $"Document Number:\n{model.FullDocumentNumber}", $"Revision Number:\n{rev}" },
            new[] { $"Title: {model.Title}", "", "Page Number:" }
        }, firstRowBold: true);

        var confidentiality = TextParagraph(ConfidentialityStatement, bold: false, halfPointSize: 16); // 8pt

        AddHeaderToAllPages(doc, identity, confidentiality);
    }

    // Body = the eight sections in order; Revision History (8) is a table.
    private static void BuildBody(Body body, SopDocument model, ConversionOptions options)
    {
        foreach (TargetSection ts in Enum.GetValues<TargetSection>())
        {
            AppendToBody(body, TextParagraph($"{(int)ts}. {ts.ToString().ToUpperInvariant()}:", bold: true));

            if (ts == TargetSection.RevisionHistory)
            {
                AppendToBody(body, BuildRevisionTable(model, options));
                continue;
            }

            var section = model.Sections[ts];
            if (section.IsEmpty)
                AppendToBody(body, TextParagraph("N/A"));
            else
                foreach (var line in section.Paragraphs)
                    AppendToBody(body, TextParagraph(line));
        }
    }

    private static Table BuildRevisionTable(SopDocument model, ConversionOptions options)
    {
        var rows = new List<string[]> { new[] { "Revision Number:", "Effective Date:", "Reason for Revision:" } };
        foreach (var e in model.RevisionHistory.TakeLast(Math.Max(0, options.PriorRevisionsToKeep)))
            rows.Add(new[] { e.RevisionNumber, e.EffectiveDate, e.Reason });

        var newRev = HeaderBuilder.FormatRevision(model.NewRevisionNumber, options.RevisionPadding);
        rows.Add(new[] { newRev, "(see header)", options.NewRevisionReason });

        return BorderedTable(rows, firstRowBold: true);
    }
}
