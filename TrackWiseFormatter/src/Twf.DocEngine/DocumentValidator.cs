using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace Twf.DocEngine;

/// <summary>
/// A quality gate: validates a generated document against the OOXML schema. This catches
/// structural defects (e.g. a table with no &lt;w:tblGrid&gt;) before the file ever reaches
/// Word, and produces validation evidence useful for deployment sign-off (SOW D3.2).
/// </summary>
public static class DocumentValidator
{
    public static IReadOnlyList<string> Validate(string path)
    {
        using var doc = WordprocessingDocument.Open(path, isEditable: false);
        var validator = new OpenXmlValidator(FileFormatVersions.Office2019);

        return validator.Validate(doc)
            .Select(e => $"{e.Description} [part: {e.Part?.Uri}; xpath: {e.Path?.XPath}]")
            .ToList();
    }
}
