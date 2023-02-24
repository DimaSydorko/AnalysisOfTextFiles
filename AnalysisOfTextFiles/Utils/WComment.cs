using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AnalysisOfTextFiles.Objects;

public class WComment
{
  public static void Add(MainDocumentPart mainPart, Paragraph paragraph, string message){
    int id = 0;
    Comments comments;
    
    // Verify that the document contains a
    // WordProcessingCommentsPart part; if not, add a new one.
    if (mainPart.GetPartsCountOfType<WordprocessingCommentsPart>() > 0)
    {
      comments =
        mainPart.WordprocessingCommentsPart.Comments;
      if (comments.HasChildren)
      {
        // Obtain an unused ID.
        id = Int32.Parse(comments.Descendants<Comment>().Select(e => e.Id.Value).Max()) + 1;
      }
    }
    else
    {
      // No WordprocessingCommentsPart part exists, so add one to the package.
      WordprocessingCommentsPart commentPart =
        mainPart.AddNewPart<WordprocessingCommentsPart>();
      commentPart.Comments = new Comments();
      comments = commentPart.Comments;
    }

    // Compose a new Comment and add it to the Comments part.
    Paragraph par = new Paragraph(new Run(new Text(message)));
    Comment cmt =
      new Comment()
      {
        Id = id.ToString(),
        Author = "Wrong style",
      };
    cmt.AppendChild(par);
    comments.AppendChild(cmt);
    comments.Save();

    // Specify the text range for the Comment.
    // Insert the new CommentRangeStart before the first run of paragraph.
    paragraph.InsertBefore(new CommentRangeStart() { Id = id.ToString() }, paragraph.GetFirstChild<Run>());

    // Insert the new CommentRangeEnd after last run of paragraph.
    CommentRangeEnd cmtEnd = paragraph.InsertAfter(new CommentRangeEnd() { Id = id.ToString() }, paragraph.Elements<Run>().Last());

    // Compose a run with CommentReference and insert it.
    paragraph.InsertAfter(new Run(new CommentReference() { Id = id.ToString() }), cmtEnd);
}

}