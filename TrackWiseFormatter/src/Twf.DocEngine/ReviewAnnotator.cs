using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Twf.DocEngine;

/// <summary>
/// Adds Word review comments in the margin, anchored to a specific paragraph. This turns the
/// Exception Report into in-context annotations a reviewer sees exactly where the issue is
/// (SOW §4.2 "make uncertain conditions visible"). Comments live in a WordprocessingCommentsPart
/// and are referenced from the body via CommentRangeStart/End + CommentReference.
/// </summary>
public sealed class ReviewAnnotator
{
    private readonly string _author;
    private int _nextId = 1;

    public ReviewAnnotator(string author = "TrackWise Formatter") => _author = author;

    public void AddComment(WordprocessingDocument doc, Paragraph target, string message, System.DateTime dateUtc)
    {
        var main = doc.MainDocumentPart!;
        var part = main.GetPartsOfType<WordprocessingCommentsPart>().FirstOrDefault();
        if (part is null)
        {
            part = main.AddNewPart<WordprocessingCommentsPart>();
            part.Comments = new Comments();
        }

        var id = (_nextId++).ToString();

        part.Comments.Append(new Comment(
            new Paragraph(new Run(new Text(message) { Space = SpaceProcessingModeValues.Preserve })))
        {
            Id = id,
            Author = _author,
            Initials = "TWF",
            Date = dateUtc
        });

        // Anchor the comment across the paragraph. CommentRangeStart must come after pPr.
        var start = new CommentRangeStart { Id = id };
        var pPr = target.GetFirstChild<ParagraphProperties>();
        if (pPr is not null) target.InsertAfter(start, pPr);
        else target.InsertAt(start, 0);

        target.Append(new CommentRangeEnd { Id = id });
        target.Append(new Run(new CommentReference { Id = id }));
    }
}
