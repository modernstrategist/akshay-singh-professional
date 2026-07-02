using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Twf.Core;
using static Twf.DocEngine.OpenXmlHelpers;

namespace Twf.DocEngine;

/// <summary>
/// Clears the template body and writes target sections 1-7 (Purpose..Forms) with the
/// mapped legacy content, emitting "N/A" for any section with no source content.
/// Section 8 (Revision History) is written by <see cref="RevisionHistoryBuilder"/>.
/// The body-level &lt;w:sectPr&gt; is preserved and kept last.
/// </summary>
public sealed class SectionBuilder
{
    public void Apply(Body body, SopDocument model)
    {
        // Remove existing content but keep the section properties (page size/margins/header refs).
        foreach (var e in body.Elements<Paragraph>().ToList()) e.Remove();
        foreach (var e in body.Elements<Table>().ToList()) e.Remove();

        var sectPr = body.Elements<SectionProperties>().LastOrDefault();

        foreach (TargetSection ts in Enum.GetValues<TargetSection>())
        {
            if (ts == TargetSection.RevisionHistory) continue; // handled by RevisionHistoryBuilder

            Insert(body, sectPr, Para($"{(int)ts}. {ts.ToString().ToUpperInvariant()}:", bold: true));

            var section = model.Sections[ts];
            if (section.IsEmpty)
                Insert(body, sectPr, Para("N/A"));
            else
                foreach (var line in section.Paragraphs)
                    Insert(body, sectPr, Para(line));
        }
    }

    internal static void Insert(Body body, SectionProperties? sectPr, OpenXmlElement element)
    {
        if (sectPr is not null) body.InsertBefore(element, sectPr);
        else body.Append(element);
    }
}
