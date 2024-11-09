using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;

namespace AnalysisOfTextFiles.Objects;

public class WReport
{
  public enum TitleType
  {
    Empty,
    Order,
    Edited,
    Wrong
  }  
  public enum OrderType
  {
    MissedAfter,
    MissedBefore,
    InsteadBefore,
    InsteadAfter,
  }
  public static void Write(string reportMessage, bool isRewrite = false, bool isSettings = false)
  {
    var mess = $"{reportMessage}\n";
    var filePath = isSettings ? State.FilePath.stylesSettings : State.FilePath.report;
    
    if (isRewrite) File.WriteAllText(filePath, mess);
    else File.AppendAllText(filePath, mess);
  }

  public static void CreateReportFile()
  {
    var timestamp = DateTime.Now.ToString("F");
    Write($"-----------------Report ({timestamp})----------------", true);
  }
    public static void CreateStylesFile()
  {
    var timestamp = DateTime.Now.ToString("F");
    Write($"-----------------Styles settings ({timestamp})----------------", true, true);
  }
  
  public static void Order(Paragraph paragraph, CheckParagraph.ContentType type, int idx, string styleName, List<string> order, OrderType orderType, string? wrongOrder = "",  WTable? table = null)
  {
    string allowed = string.Join(", ", order.Select(e => $"'{e}'"));
    string extraMessage = orderType switch
    {
      OrderType.MissedAfter => $"mast be after {allowed}",
      OrderType.MissedBefore => $"mast be before {allowed}",
      OrderType.InsteadAfter => $"mast be after {allowed}, But now is '{wrongOrder}'",
      OrderType.InsteadBefore => $"mast be before {allowed}, But now is '{wrongOrder}'",
      _ => $"mast be after {allowed}"
    };
    
    OnMessage(paragraph, type, idx, styleName, table, TitleType.Order, extraMessage);
  }

  public static void OnMessage(
    Paragraph paragraph,
    CheckParagraph.ContentType type,
    int idx,
    string styleName,
    WTable? table = null,
    TitleType? title = TitleType.Wrong,
    string? extraMessage = null)
  {
    string message = title switch
    {
      TitleType.Empty => "Empty Line",
      TitleType.Wrong => $"Style: '{styleName}'",
      TitleType.Order => $"Invalid order, style: '{styleName}' {extraMessage}",
      TitleType.Edited => $"Edited style: '{styleName}'",
      _ => $"Style: '{styleName}'"
    };
    
    bool isComment = type != CheckParagraph.ContentType.Header && type != CheckParagraph.ContentType.Footer;
    if (isComment && State.IsСomments)
    {
      _AddComment(paragraph, message, title);
    }

    if (title != TitleType.Empty)
    {
      string text = paragraph.InnerText;
      string firstLetters = text.Length > 15 ? text.Substring(0, 15) + "..." : text;

      if (firstLetters.Length > 0)
      {
        message = $"('{firstLetters}') {message}";
      }
    }
    
    string report = type switch
    {
      CheckParagraph.ContentType.Paragraph => $"{type} {idx + 1} {message}",
      CheckParagraph.ContentType.TOC => $"TOC Paragraph {idx + 1} {message}",
      CheckParagraph.ContentType.Table => $"Table {table?.Idx + 1}, Row {table?.RowIdx + 1}, Cell {table?.CellIdx + 1}, Par {table?.ParIdx + 1} {message}",
      _ => $"{type} {idx + 1} {message}"
    };

    Write(report);
  }

  private static void _AddComment(Paragraph paragraph, string message, TitleType? title = TitleType.Wrong)
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

    string Author = title switch
    {
      TitleType.Empty => "Empty Line",
      TitleType.Wrong => "Wrong style",
      TitleType.Edited => "Edited style",
      TitleType.Order => "Invalid Order",
      _ => string.Empty
    };
    
    // Compose a new Comment and add it to the Comments part.
    var par = new Paragraph(new Run(new Text(message)));
    var cmt =
      new Comment
      {
        Id = id.ToString(),
        Author = Author
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

  public static string OnCompareStyleSettings(object settings, object value)
  {
    var type = settings.GetType();
    var properties = type.GetProperties();
    var diff = "";

    foreach (var property in properties)
    {
      if (property.Name != "after" && property.Name != "before")
      {
        var setting = property.GetValue(settings);
        var currValue = property.GetValue(value);

        if (!Equals(setting, currValue))
        {
          diff += $"{property.Name}: {setting} -> {currValue}\n";
        }
      }
    }

    return diff;
  }
}