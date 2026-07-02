namespace Twf.Core;

/// <summary>Legacy SOP source family. DDR uses the older header; NTR is closer to target (Phase 2).</summary>
public enum DocType { Ddr, Ntr, Unknown }

/// <summary>The eight mandatory TrackWise SOP sections, in their canonical order.</summary>
public enum TargetSection
{
    Purpose = 1,
    Scope = 2,
    Responsibilities = 3,
    References = 4,
    Definitions = 5,
    Procedure = 6,
    Forms = 7,
    RevisionHistory = 8
}

public enum Severity { Info, Warning, Error }

/// <summary>A single entry in the Exception Report (missing field, failed lookup, deferred rule, etc.).</summary>
public sealed record ExceptionRecord(Severity Severity, string Type, string Message, string? Section = null);

/// <summary>One target section plus the source content mapped into it.</summary>
public sealed class SopSection
{
    public required TargetSection Target { get; init; }
    public List<string> Paragraphs { get; } = new();
    public bool IsEmpty => Paragraphs.Count == 0;
}

/// <summary>A row of the Revision History table.</summary>
public sealed class RevisionEntry
{
    public string RevisionNumber { get; init; } = "";
    public string EffectiveDate { get; init; } = "";
    public string Reason { get; init; } = "";
}

/// <summary>The internal model the parser produces and the builders consume.</summary>
public sealed class SopDocument
{
    public string? LegacyId { get; set; }                 // e.g. GEN-027
    public string? FullDocumentNumber { get; set; }       // e.g. SOP-DDR-GEN-000003
    public string Title { get; set; } = "Untitled";
    public string Department { get; set; } = "General";
    public DocType DocType { get; set; } = DocType.Unknown;

    public int OldRevisionNumber { get; set; }
    public int NewRevisionNumber => OldRevisionNumber + 1;

    public Dictionary<TargetSection, SopSection> Sections { get; } = new();
    public List<RevisionEntry> RevisionHistory { get; } = new();
    public List<ExceptionRecord> Exceptions { get; } = new();
}
