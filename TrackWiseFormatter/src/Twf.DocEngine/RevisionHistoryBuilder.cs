using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Twf.Core;
using static Twf.DocEngine.OpenXmlHelpers;

namespace Twf.DocEngine;

/// <summary>
/// Writes section 8 (Revision History) as a four-line rolling table:
/// the last N prior revisions plus a new row carrying the incremented revision number,
/// "(see header)" as the effective date, and the configured reason.
/// </summary>
public sealed class RevisionHistoryBuilder
{
    public void Apply(Body body, SopDocument model, ConversionOptions options)
    {
        var sectPr = body.Elements<SectionProperties>().LastOrDefault();

        SectionBuilder.Insert(body, sectPr,
            Para($"{(int)TargetSection.RevisionHistory}. REVISION HISTORY:", bold: true));

        var prior = model.RevisionHistory
            .TakeLast(Math.Max(0, options.PriorRevisionsToKeep))
            .ToList();
        var newRev = HeaderBuilder.FormatRevision(model.NewRevisionNumber, options.RevisionPadding);

        var table = new Table(BuildTableProperties());
        table.Append(Row(bold: true, "Revision Number:", "Effective Date:", "Reason for Revision:"));
        foreach (var e in prior)
            table.Append(Row(false, e.RevisionNumber, e.EffectiveDate, e.Reason));
        table.Append(Row(false, newRev, "(see header)", options.NewRevisionReason));

        SectionBuilder.Insert(body, sectPr, table);
    }

    private static TableProperties BuildTableProperties() => new(
        new TableBorders(
            new TopBorder { Val = BorderValues.Single, Size = 4U },
            new BottomBorder { Val = BorderValues.Single, Size = 4U },
            new LeftBorder { Val = BorderValues.Single, Size = 4U },
            new RightBorder { Val = BorderValues.Single, Size = 4U },
            new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4U },
            new InsideVerticalBorder { Val = BorderValues.Single, Size = 4U }),
        new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" });

    private static TableRow Row(bool bold, params string[] cells)
    {
        var tr = new TableRow();
        foreach (var c in cells)
            tr.Append(new TableCell(Para(c, bold)));
        return tr;
    }
}
