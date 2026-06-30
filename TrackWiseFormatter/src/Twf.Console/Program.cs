using Twf.Core;
using Twf.DocEngine;

// Usage: twf <source.docx> <template.docx> <output.docx> [legacy-ids.csv]
if (args.Length < 3)
{
    Console.WriteLine("TrackWise Formatting POC (Pilot / Phase 1)");
    Console.WriteLine("Usage: twf <source.docx> <template.docx> <output.docx> [legacy-ids.csv]");
    return 1;
}

string source = args[0], template = args[1], output = args[2];

// Legacy-ID map: from CSV if supplied, otherwise the GEN-027 pilot seed.
LegacyIdMap legacy = args.Length >= 4 && File.Exists(args[3])
    ? LegacyIdMap.FromCsv(args[3])
    : new LegacyIdMap().Add("GEN-027", "SOP-DDR-GEN-000003", "Handling of Manufacturing Reworks");

var options = new ConversionOptions
{
    RevisionPadding = 3,   // matches the approved template; set 0 to honour SOW "omit leading zeroes"
    TitleCase = true
};

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
