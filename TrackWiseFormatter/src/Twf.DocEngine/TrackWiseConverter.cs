using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using Twf.Core;

namespace Twf.DocEngine;

/// <summary>
/// Orchestrates a single conversion: parse the legacy SOP, resolve its identity via the
/// legacy-ID map, copy the approved template, then apply the header / section /
/// revision-history builders against it. The source file is never modified.
/// </summary>
public sealed class TrackWiseConverter
{
    private readonly DocxParser _parser = new();
    private readonly HeaderBuilder _header = new();
    private readonly SectionBuilder _sections = new();
    private readonly RevisionHistoryBuilder _revisions = new();
    private readonly LegacyIdMap _legacy;
    private readonly ConversionOptions _options;

    public TrackWiseConverter(LegacyIdMap legacy, ConversionOptions? options = null)
    {
        _legacy = legacy;
        _options = options ?? new ConversionOptions();
    }

    public ConversionResult Convert(string sourcePath, string templatePath, string outputPath,
        DocType docType = DocType.Ddr)
    {
        var model = _parser.Parse(sourcePath, docType);
        ResolveIdentity(model);

        // Build against a copy of the approved template (authentic styles + confidentiality header).
        File.Copy(templatePath, outputPath, overwrite: true);
        using (var doc = WordprocessingDocument.Open(outputPath, isEditable: true))
        {
            var body = doc.MainDocumentPart!.Document.Body!;
            _header.Apply(doc, model, _options);
            _sections.Apply(body, model);          // sections 1-7
            _revisions.Apply(body, model, _options); // section 8
            doc.MainDocumentPart.Document.Save();
        }

        AddAdvisoryExceptions(model);
        return new ConversionResult { OutputPath = outputPath, Document = model };
    }

    private void ResolveIdentity(SopDocument model)
    {
        if (model.LegacyId is { Length: > 0 } legacy)
        {
            if (_legacy.TryResolve(legacy, out var full, out var mappedTitle))
            {
                model.FullDocumentNumber = full;
                if (mappedTitle is { Length: > 0 }) model.Title = mappedTitle;
            }
            else
            {
                model.FullDocumentNumber = legacy;
                model.Exceptions.Add(new ExceptionRecord(Severity.Error, "LegacyLookup",
                    $"Legacy ID '{legacy}' not found in mapping; kept legacy number.", "Header"));
            }
        }
        else
        {
            model.Exceptions.Add(new ExceptionRecord(Severity.Warning, "LegacyLookup",
                "No legacy SOP number found in source header.", "Header"));
        }

        if (_options.TitleCase)
            model.Title = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(model.Title.ToLowerInvariant());
    }

    private static void AddAdvisoryExceptions(SopDocument model)
    {
        if (model.Sections[TargetSection.Forms].IsEmpty)
            model.Exceptions.Add(new ExceptionRecord(Severity.Warning, "MissingSection",
                "Source has no Forms/Attachments section; set to N/A.", "Forms"));

        model.Exceptions.Add(new ExceptionRecord(Severity.Info, "ImageCarryover",
            "Embedded exhibits/flowchart images are not carried into the draft in Phase 1; flagged for manual placement (Phase 3).", "Procedure"));

        model.Exceptions.Add(new ExceptionRecord(Severity.Info, "Numbering",
            "Procedure sub-step numbering normalization (e.g. I.A.1 -> 1.1.1) is deferred to Phase 3.", "Procedure"));

        if (model.OldRevisionNumber == 0)
            model.Exceptions.Add(new ExceptionRecord(Severity.Warning, "RevisionParse",
                "Could not read old revision number from header; increment defaulted from 0.", "Header"));
    }
}
