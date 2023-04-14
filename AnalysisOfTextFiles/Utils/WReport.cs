using System;
using System.IO;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AnalysisOfTextFiles.Objects;

using ContentType = Analis.ContentType;

public class WReport
{
  public static void CreateReportFile()
  {
    string timestamp = DateTime.Now.ToString("F");
    File.WriteAllText(State.FilePath.report, $"-----------------Report ({timestamp})----------------\n");
  }

  public static void OnMessage(Paragraph paragraph, ContentType type, int idx, string styleName, WTable? table = null)
  {
    string text = paragraph.InnerText;
    string firstLetters = text.Length > 15 ? text.Substring(0, 15) + "..." : text;

    bool isComment = type != ContentType.Header || type != ContentType.Footer;
    if (isComment && State.IsСomments)
    {
      _AddComment(paragraph, styleName);
    }

    string parData = $" ('{firstLetters}') Style: {styleName}\n";
    string report = $"{type} {idx + 1} {parData}";

    //TOC: Table of content
    if (type == ContentType.TOC) report = $"TOC Paragraph {idx + 1} {parData}";
    else if (table != null)
      report =
        $"Table {table.Idx + 1}, Row {table.RowIdx + 1}, Cell {table.CellIdx + 1}, Par {table.ParIdx + 1} {parData}";
    
    File.AppendAllText(State.FilePath.report, report);
  }

  private static void _AddComment(Paragraph paragraph, string message)
  {
    int id = 0;
    Comments comments;
    MainDocumentPart mainPart = State.WDocument.MainDocumentPart;

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
    CommentRangeEnd cmtEnd = paragraph.InsertAfter(new CommentRangeEnd() { Id = id.ToString() },
      paragraph.Elements<Run>().LastOrDefault());

    // Compose a run with CommentReference and insert it.
    paragraph.InsertAfter(new Run(new CommentReference() { Id = id.ToString() }), cmtEnd);
  }

  public static string OnCompareObjects(object obj1, object obj2)
  {
    Type type = obj1.GetType();
    PropertyInfo[] properties = type.GetProperties();
    string diff = "";

    foreach (PropertyInfo property in properties)
    {
      object value1 = property.GetValue(obj1);
      object value2 = property.GetValue(obj2);

      if (!Equals(value1, value2))
      {
        diff += $"{property.Name}: {value1} -> {value2}\n";
      }
    }

    return diff;
  }
}