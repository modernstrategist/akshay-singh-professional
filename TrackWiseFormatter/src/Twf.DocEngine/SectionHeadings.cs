using System.Text.RegularExpressions;
using Twf.Core;

namespace Twf.DocEngine;

/// <summary>
/// Shared logic for recognising SOP section headings, used both when reading the source
/// (<see cref="DocxParser"/>) and when locating headings in the output template
/// (<see cref="TemplateFiller"/>). One definition, so both sides agree.
/// </summary>
public static class SectionHeadings
{
    public static readonly IReadOnlyDictionary<string, TargetSection> Aliases =
        new Dictionary<string, TargetSection>(StringComparer.OrdinalIgnoreCase)
        {
            ["PURPOSE"] = TargetSection.Purpose,
            ["SCOPE"] = TargetSection.Scope,
            ["RESPONSIBILITY"] = TargetSection.Responsibilities,
            ["RESPONSIBILITIES"] = TargetSection.Responsibilities,
            ["REFERENCE"] = TargetSection.References,
            ["REFERENCES"] = TargetSection.References,
            ["DEFINITION"] = TargetSection.Definitions,
            ["DEFINITIONS"] = TargetSection.Definitions,
            ["PROCEDURE"] = TargetSection.Procedure,
            ["ATTACHMENT"] = TargetSection.Forms,
            ["ATTACHMENTS"] = TargetSection.Forms,
            ["FORM"] = TargetSection.Forms,
            ["FORMS"] = TargetSection.Forms,
            ["RECORD OF REVISIONS"] = TargetSection.RevisionHistory,
            ["REVISION HISTORY"] = TargetSection.RevisionHistory,
        };

    /// <summary>Uppercase, drop a trailing colon, strip any parenthetical, collapse spaces.</summary>
    public static string Normalize(string text)
    {
        var key = text.Trim().ToUpperInvariant().TrimEnd(':').Trim();
        key = Regex.Replace(key, @"\(.*?\)", "").Trim();      // e.g. "(continued)", "(include only 4 revisions)"
        return Regex.Replace(key, @"\s+", " ");
    }

    public static bool TryMatch(string text, out TargetSection target)
        => Aliases.TryGetValue(Normalize(text), out target);
}
