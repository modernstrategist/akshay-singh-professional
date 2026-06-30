namespace Twf.Core;

/// <summary>
/// Tunable formatting rules. The defaults encode the decisions we recommend pinning
/// in the Phase 1 specification (D1.4). Two of these are OPEN DECISIONS flagged in the SOW.
/// </summary>
public sealed class ConversionOptions
{
    /// <summary>
    /// Leading-zero width for the revision number in the header and history table.
    /// The approved template shows 3 digits ("006"); the SOW Section 5.2 says
    /// "omit leading zeroes". OPEN DECISION — set to 0 for no padding.
    /// </summary>
    public int RevisionPadding { get; set; } = 3;

    /// <summary>
    /// Apply Title Case to the document title. The legacy source is ALL CAPS; the
    /// template uses Title Case. OPEN DECISION.
    /// </summary>
    public bool TitleCase { get; set; } = true;

    /// <summary>Reason text written on the newly added revision-history row.</summary>
    public string NewRevisionReason { get; set; } = "Updated to TrackWise format.";

    /// <summary>How many prior revision rows to retain (the rule is a 4-line rolling table = 3 prior + 1 new).</summary>
    public int PriorRevisionsToKeep { get; set; } = 3;
}
