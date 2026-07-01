using Twf.Core;
using Twf.DocEngine;
using Twf.DocEngine.Learning;

// Two modes:
//   Convert:  twf <source.docx> <template.docx> <output.docx> [legacy-ids.csv]
//   Learn:    twf --scratch [output.docx]     (builds a sample doc FROM SCRATCH, no inputs)
if (args.Length >= 1 && args[0].Equals("--scratch", StringComparison.OrdinalIgnoreCase))
{
    var outPath = args.Length >= 2 ? args[1] : "scratch-demo.docx";
    ScratchDocumentWriter.Write(DemoModel(), new ConversionOptions(), outPath);
    Console.WriteLine($"Built a TrackWise SOP from scratch (no template): {outPath}");
    Console.WriteLine("Open it in Word, then rename it to .zip and read word/document.xml to see the tree.");
    return 0;
}

if (args.Length < 3)
{
    Console.WriteLine("TrackWise Formatting POC (Pilot / Phase 1)");
    Console.WriteLine("  Convert : twf <source.docx> <template.docx> <output.docx> [legacy-ids.csv]");
    Console.WriteLine("  Learn   : twf --scratch [output.docx]");
    return 1;
}

string source = args[0], template = args[1], output = args[2];

LegacyIdMap legacy = args.Length >= 4 && File.Exists(args[3])
    ? LegacyIdMap.FromCsv(args[3])
    : new LegacyIdMap().Add("GEN-027", "SOP-DDR-GEN-000003", "Handling of Manufacturing Reworks");

var options = new ConversionOptions { RevisionPadding = 3, TitleCase = true };
var converter = new TrackWiseConverter(legacy, options);
ConversionResult result = converter.Convert(source, template, output, DocType.Ddr);

var d = result.Document;
Console.WriteLine($"Converted : {d.LegacyId} -> {d.FullDocumentNumber}");
Console.WriteLine($"Title     : {d.Title}");
Console.WriteLine($"Revision  : {d.OldRevisionNumber} -> {d.NewRevisionNumber}");
Console.WriteLine($"Sections  : {string.Join(", ", d.Sections.Values.Select(s => $"{(int)s.Target}:{(s.IsEmpty ? "N/A" : s.Paragraphs.Count.ToString())}"))}");
Console.WriteLine($"Output    : {result.OutputPath}");
Console.WriteLine();
Console.WriteLine("Exception Report:");
foreach (var e in result.Exceptions)
    Console.WriteLine($"  [{e.Severity}] {e.Type} ({e.Section}): {e.Message}");

return result.HasErrors ? 2 : 0;

// A small in-memory sample so the --scratch demo needs no client files.
static SopDocument DemoModel()
{
    var m = new SopDocument
    {
        LegacyId = "GEN-027",
        FullDocumentNumber = "SOP-DDR-GEN-000003",
        Title = "Handling Of Manufacturing Reworks",
        Department = "General",
        DocType = DocType.Ddr,
        OldRevisionNumber = 5
    };
    foreach (TargetSection ts in Enum.GetValues<TargetSection>())
        m.Sections[ts] = new SopSection { Target = ts };

    m.Sections[TargetSection.Purpose].Paragraphs.Add("The purpose of this procedure is to define the process flow for materials that have the potential to be reclaimed through the rework process.");
    m.Sections[TargetSection.Scope].Paragraphs.Add("This procedure applies to Dietary Supplements manufactured at the IVC facility.");
    m.Sections[TargetSection.Responsibilities].Paragraphs.Add("QC Supervisor or Designee will notify the Distressed Batch Team of potential reworks.");
    m.Sections[TargetSection.References].Paragraphs.Add("QCLAB-007 Stability Program: Operation of Program");
    m.Sections[TargetSection.Definitions].Paragraphs.Add("DAF - Destruction Authorization Form.");
    m.Sections[TargetSection.Procedure].Paragraphs.Add("Notify the Distressed Batch Team of a lot (batch) that has potential to be reworked.");
    // Forms deliberately left empty -> renders as N/A

    m.RevisionHistory.Add(new RevisionEntry { RevisionNumber = "3", EffectiveDate = "11/23/15", Reason = "Replaced old logo; corrected job titles." });
    m.RevisionHistory.Add(new RevisionEntry { RevisionNumber = "4", EffectiveDate = "11/21/17", Reason = "Updated to IVC format." });
    m.RevisionHistory.Add(new RevisionEntry { RevisionNumber = "5", EffectiveDate = "12/21/20", Reason = "Updated job titles." });
    return m;
}
