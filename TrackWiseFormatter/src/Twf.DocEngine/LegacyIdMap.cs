using Twf.Core;

namespace Twf.DocEngine;

/// <summary>
/// Maps a legacy document ID (e.g. GEN-027) to its Full Document Number
/// (e.g. SOP-DDR-GEN-000003) and optional canonical title.
/// In Phase 2 this is backed by the IVC "Legacy Document IDs" spreadsheet imported into SQL;
/// here it is a simple in-memory / CSV lookup. Failed lookups are reported as exceptions.
/// </summary>
public sealed class LegacyIdMap
{
    private readonly Dictionary<string, (string Full, string? Title)> _map =
        new(StringComparer.OrdinalIgnoreCase);

    public LegacyIdMap Add(string legacyId, string fullNumber, string? title = null)
    {
        _map[legacyId] = (fullNumber, title);
        return this;
    }

    /// <summary>All legacy IDs known to the map (used by the reference rewriter).</summary>
    public IEnumerable<string> KnownIds => _map.Keys;

    /// <summary>Load from a CSV with header row: LegacyId,FullDocumentNumber,Title</summary>
    public static LegacyIdMap FromCsv(string path)
    {
        var map = new LegacyIdMap();
        foreach (var line in File.ReadLines(path).Skip(1))
        {
            if (string.IsNullOrWhiteSpace(line)) continue;
            var parts = line.Split(',');
            if (parts.Length >= 2)
                map.Add(parts[0].Trim(), parts[1].Trim(), parts.Length >= 3 ? parts[2].Trim() : null);
        }
        return map;
    }

    public bool TryResolve(string legacyId, out string fullNumber, out string? title)
    {
        if (_map.TryGetValue(legacyId, out var v))
        {
            fullNumber = v.Full;
            title = v.Title;
            return true;
        }
        fullNumber = legacyId;
        title = null;
        return false;
    }
}
