using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Twf.DocEngine;

/// <summary>Low-level WordprocessingML helpers shared by the builders.</summary>
internal static class OpenXmlHelpers
{
    /// <summary>Create a body paragraph carrying a single run of text.</summary>
    public static Paragraph Para(string text, bool bold = false, bool highlightYellow = false)
    {
        var run = new Run();
        var rpr = new RunProperties();
        if (bold) rpr.Append(new Bold());
        if (highlightYellow) rpr.Append(new Highlight { Val = HighlightColorValues.Yellow });
        if (rpr.HasChildren) run.Append(rpr);
        run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        return new Paragraph(run);
    }

    /// <summary>
    /// Replace a table cell's text, keeping its TableCellProperties. Supports multi-line
    /// values via '\n' (rendered as line breaks). The first line is bold (the label).
    /// </summary>
    public static void WriteCell(TableCell cell, string text)
    {
        foreach (var p in cell.Elements<Paragraph>().ToList())
            p.Remove();

        var para = new Paragraph();
        var lines = text.Split('\n');
        for (int i = 0; i < lines.Length; i++)
        {
            var run = new Run();
            if (i == 0) run.Append(new RunProperties(new Bold()));
            if (i > 0) run.Append(new Break());
            run.Append(new Text(lines[i]) { Space = SpaceProcessingModeValues.Preserve });
            para.Append(run);
        }
        cell.Append(para);
    }

    public static string CellText(TableCell cell) => cell.InnerText.Trim();

    /// <summary>
    /// Reconstruct visible lines from a header/footer part, expanding paragraph and
    /// &lt;w:br/&gt; boundaries that InnerText would otherwise collapse.
    /// </summary>
    public static IEnumerable<string> VisibleLines(OpenXmlElement root)
    {
        foreach (var para in root.Descendants<Paragraph>())
        {
            var sb = new System.Text.StringBuilder();
            foreach (var node in para.Descendants())
            {
                switch (node)
                {
                    case Text t: sb.Append(t.Text); break;
                    case Break: yield return sb.ToString(); sb.Clear(); break;
                }
            }
            yield return sb.ToString();
        }
    }
}
