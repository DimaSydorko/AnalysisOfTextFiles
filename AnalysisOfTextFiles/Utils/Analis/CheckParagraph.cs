using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AnalysisOfTextFiles.Objects;

public class CheckParagraph
{
  public enum ContentType
  {
    Header,
    Paragraph,
    Table,
    Footer,
    TOC
  }
  
  public static List<string> allowedStyles = new() { "Heading 1", "Heading 2", "Heading 3", "TOC1", "TOC2" };

  public static bool IsValidStyle(string styleName)
  {
    var keyWordLength = State.KeyWord.Length;
    if (keyWordLength <= styleName.Length)
    {
      var firstLetters = styleName.Substring(0, keyWordLength);
      return allowedStyles.Contains(styleName) || firstLetters == State.KeyWord;
    }

    return false;
  }

  public static string GetParagraphStyle(Paragraph? paragraph)
  {
    if (paragraph == null) return "Unknown";

    bool isParExist = paragraph.ParagraphProperties != null &&
                      paragraph.ParagraphProperties?.ParagraphStyleId?.Val != null;

    if (isParExist)
    {
      string? styleName = paragraph.ParagraphProperties?.ParagraphStyleId?.Val;
      var style = WStyle.GetDecodedStyle(styleName);
      return style;
    }

    return "Normal";
  }

  public static void ParagraphCheck(Paragraph paragraph, int idx, ContentType type, WTable? table = null)
  {
    var isParaExist = paragraph.ParagraphProperties != null;
    var isInnerText = !string.IsNullOrEmpty(paragraph.InnerText);

    void onComment(string styleName, WReport.TitleType? titleType = WReport.TitleType.Wrong)
    {
      WReport.OnMessage(paragraph, type, idx, styleName, table, titleType);
    }

    if (isInnerText)
    {
      if (isParaExist && paragraph.ParagraphProperties?.ParagraphStyleId != null)
      {
        string? styleName = WDecoding.RemoveSuffixIfExists(GetParagraphStyle(paragraph));
        Order.CheckParagraph(paragraph, type, styleName, idx);

        // if the value of the pStyle is allowed => skip the paragraph
        if (styleName != null)
        {
          if (!IsValidStyle(styleName)) onComment(styleName);
          else if (CheckEdited.IsEditedStyle(paragraph)) onComment(styleName, WReport.TitleType.Edited);
        }
        else
        {
          var dec = WDecoding.GetOldDecStyle(styleName);
          if (!allowedStyles.Contains(dec))
          {
            if (dec == null) onComment($"Undefined Style name '{styleName}'");
            else onComment(dec);
          }
        }
      }
      else if (isInnerText)
      {
        onComment("Normal");
      }
    }
    else
    {
      WReport.OnMessage(paragraph, type, idx, "", table, WReport.TitleType.Empty);
    }
  }
}