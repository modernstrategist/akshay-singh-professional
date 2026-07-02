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
    private readonly TemplateFiller _filler = new();
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

        // A single timestamp for all review comments in this run (caller-independent).
        var stamp = System.DateTime.UtcNow;

        // Build against a copy of the approved template (authentic styles + confidentiality header).
        File.Copy(templatePath, outputPath, overwrite: true);
        using (var doc = WordprocessingDocument.Open(outputPath, isEditable: true))
        {
            var body = doc.MainDocumentPart!.Document.Body!;
            var annotator = _options.EnableReviewComments ? new ReviewAnnotator(_options.ReviewAuthor) : null;

            _header.Apply(doc, model, _options);
            _filler.Fill(doc, model, _options, annotator, stamp);   // content under template headings + table

            if (_options.EnableReferenceRewrite)
                new ReferenceRewriter(_legacy).Rewrite(body, model);

            AddAdvisoryExceptions(model);
            AnnotateAdvisories(doc, body, model, annotator, stamp);

            doc.MainDocumentPart.Document.Save();
        }

        if (_options.ValidateOutput)
            foreach (var err in DocumentValidator.Validate(outputPath))
                model.Exceptions.Add(new ExceptionRecord(Severity.Error, "SchemaValidation", err));

        return new ConversionResult { OutputPath = outputPath, Document = model };
    }

    // Attach the Procedure-level advisories (images, numbering) as in-context comments.
    private static void AnnotateAdvisories(WordprocessingDocument doc, DocumentFormat.OpenXml.Wordprocessing.Body body,
        SopDocument model, ReviewAnnotator? annotator, System.DateTime stamp)
    {
        if (annotator is null) return;
        var procedure = FindHeading(body, TargetSection.Procedure);
        if (procedure is null) return;

        foreach (var ex in model.Exceptions.Where(e => e.Section == "Procedure"))
            annotator.AddComment(doc, procedure, $"[{ex.Type}] {ex.Message}", stamp);
    }

    private static DocumentFormat.OpenXml.Wordprocessing.Paragraph? FindHeading(
        DocumentFormat.OpenXml.Wordprocessing.Body body, TargetSection ts) =>
        body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault(p =>
        {
            var t = p.InnerText.Trim();
            return t.Length is > 0 and < 60 && SectionHeadings.TryMatch(t, out var m) && m == ts;
        });

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
