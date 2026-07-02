using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Wordprocessing;
using Twf.Core;

namespace Twf.DocEngine;

/// <summary>
/// Rewrites legacy document numbers found in the body to their Full Document Number, keeping
/// the legacy number in parentheses (SOW Phase 2). Uses the same <see cref="LegacyIdMap"/>.
/// Operates per &lt;w:t&gt; run so surrounding formatting is preserved; matches on whole tokens
/// so partial IDs are not replaced.
/// </summary>
public sealed class ReferenceRewriter
{
    private readonly LegacyIdMap _map;

    public ReferenceRewriter(LegacyIdMap map) => _map = map;

    /// <summary>Returns the number of replacements made, and records unresolved-lookup exceptions elsewhere.</summary>
    public int Rewrite(Body body, SopDocument model)
    {
        int replacements = 0;
        var ids = _map.KnownIds.ToList();
        if (ids.Count == 0) return 0;

        foreach (var text in body.Descendants<Text>())
        {
            var value = text.Text;
            foreach (var legacy in ids)
            {
                if (!value.Contains(legacy, System.StringComparison.OrdinalIgnoreCase)) continue;
                if (!_map.TryResolve(legacy, out var full, out _)) continue;

                var updated = ReplaceWholeToken(value, legacy, $"{full} ({legacy})");
                if (!ReferenceEquals(updated, value) && updated != value)
                {
                    value = updated;
                    replacements++;
                }
            }
            if (value != text.Text) text.Text = value;
        }
        return replacements;
    }

    // Replace the token only when it is not part of a larger alphanumeric run, and not already
    // immediately followed by " (" (which would indicate it was already rewritten).
    private static string ReplaceWholeToken(string input, string token, string replacement)
        => Regex.Replace(
            input,
            $@"(?<![A-Za-z0-9]){Regex.Escape(token)}(?!\s*\()(?![A-Za-z0-9])",
            replacement);
}
