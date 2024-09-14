using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AnalysisOfTextFiles.Objects;

public class WReport
{
  public static void Write(string reportMessage, bool isRewrite = false)
  {
    if (isRewrite)
    {
      File.WriteAllText(State.FilePath.report, $"{reportMessage}\n");
    }
    else
    {
      File.AppendAllText(State.FilePath.report, $"{reportMessage}\n");
    }
  }
  public static void CreateReportFile()
  {
    var timestamp = DateTime.Now.ToString("F");
    Write($"-----------------Report ({timestamp})----------------", true);
  }

  public static void OnMessage(Paragraph paragraph, Analis.ContentType type, int idx, string styleName, bool isEdited,
    WTable? table = null)
  {
    var text = paragraph.InnerText;
    var firstLetters = text.Length > 15 ? text.Substring(0, 15) + "..." : text;

    var isComment = type != Analis.ContentType.Header && type != Analis.ContentType.Footer;
    if (isComment && State.IsСomments)
    {
      _AddComment(paragraph, styleName);
    }

    var parData = $" ('{firstLetters}') Style: {styleName}";
    var report = $"{type} {idx + 1} {parData}";

    //TOC: Table of content
    if (type == Analis.ContentType.TOC)
    {
      report = $"TOC Paragraph {idx + 1} {parData}";
    }
    else if (table != null)
    {
      report =
        $"Table {table.Idx + 1}, Row {table.RowIdx + 1}, Cell {table.CellIdx + 1}, Par {table.ParIdx + 1} {parData}";
    }

    var isEditedText = isEdited ? "Edited " : "";

    Write($"{isEditedText}{report}");
  }

  private static void _AddComment(Paragraph paragraph, string message)
  {
    var id = 0;
    Comments comments;
    var mainPart = State.WDocument.MainDocumentPart;

    // Verify that the document contains a
    // WordProcessingCommentsPart part; if not, add a new one.
    if (mainPart.GetPartsCountOfType<WordprocessingCommentsPart>() > 0)
    {
      comments =
        mainPart.WordprocessingCommentsPart.Comments;
      if (comments.HasChildren)
      {
        // Obtain an unused ID.
        id = int.Parse(comments.Descendants<Comment>().Select(e => e.Id.Value).Max()) + 1;
      }
    }
    else
    {
      // No WordprocessingCommentsPart part exists, so add one to the package.
      var commentPart =
        mainPart.AddNewPart<WordprocessingCommentsPart>();
      commentPart.Comments = new Comments();
      comments = commentPart.Comments;
    }

    // Compose a new Comment and add it to the Comments part.
    var par = new Paragraph(new Run(new Text(message)));
    var cmt =
      new Comment
      {
        Id = id.ToString(),
        Author = "Wrong style"
      };
    cmt.AppendChild(par);
    comments.AppendChild(cmt);
    comments.Save();

    // Specify the text range for the Comment.
    // Insert the new CommentRangeStart before the first run of paragraph.
    paragraph.InsertBefore(new CommentRangeStart { Id = id.ToString() }, paragraph.GetFirstChild<Run>());

    // Insert the new CommentRangeEnd after last run of paragraph.
    var cmtEnd = paragraph.InsertAfter(new CommentRangeEnd { Id = id.ToString() },
      paragraph.Elements<Run>().LastOrDefault());

    // Compose a run with CommentReference and insert it.
    paragraph.InsertAfter(new Run(new CommentReference { Id = id.ToString() }), cmtEnd);
  }

  public static string OnCompareStyleSettings(object obj1, object obj2)
  {
    var type = obj1.GetType();
    var properties = type.GetProperties();
    var diff = "";

    foreach (var property in properties)
    {
      var value1 = property.GetValue(obj1);
      var value2 = property.GetValue(obj2);

      if (!Equals(value1, value2))
      {
        diff += $"{property.Name}: {value1} -> {value2}\n";
      }
    }

    return diff;
  }
}