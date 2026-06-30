using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Twf.Core;
using static Twf.DocEngine.OpenXmlHelpers;

namespace Twf.DocEngine;

/// <summary>
/// Populates the new TrackWise header table (Department, Document Number, Revision
/// Number, Title) on every header part. The confidentiality statement is part of the
/// approved template header and therefore already repeats on every page.
/// </summary>
public sealed class HeaderBuilder
{
    public void Apply(WordprocessingDocument doc, SopDocument model, ConversionOptions options)
    {
        var rev = FormatRevision(model.NewRevisionNumber, options.RevisionPadding);

        foreach (var hp in doc.MainDocumentPart!.HeaderParts)
        {
            foreach (var cell in hp.Header.Descendants<TableCell>())
            {
                var s = CellText(cell);
                if (s.StartsWith("Title:", StringComparison.OrdinalIgnoreCase))
                    WriteCell(cell, $"Title: {model.Title}");
                else if (s.StartsWith("Document Number:", StringComparison.OrdinalIgnoreCase))
                    WriteCell(cell, $"Document Number:\n{model.FullDocumentNumber}");
                else if (s.StartsWith("Revision Number:", StringComparison.OrdinalIgnoreCase))
                    WriteCell(cell, $"Revision Number:\n{rev}");
                else if (s.StartsWith("Department:", StringComparison.OrdinalIgnoreCase))
                    WriteCell(cell, $"Department:  {model.Department}");
            }
        }
    }

    /// <summary>Format the revision number with optional leading-zero padding (see ConversionOptions).</summary>
    public static string FormatRevision(int n, int padding)
        => padding > 0 ? n.ToString().PadLeft(padding, '0') : n.ToString();
}
