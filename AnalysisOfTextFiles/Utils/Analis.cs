using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AnalysisOfTextFiles.Objects;

public class Analis
{
  public static List<string> allowedStyles = new List<string> { "Heading1", "Heading2", "Heading3", "TOC1", "TOC2" };
  public enum ContentType
  {
    Header,
    Paragraph,
    Table,
    Footer,
    TOC
  }
  public static bool IsValidStyle(WStyle style)
  {
    int keyWordLength = State.KeyWord.Length;
    string firstLetters = style.decoded.Substring(0, keyWordLength);
    return allowedStyles.Contains(style.encoded) || firstLetters == State.KeyWord;
  }
  public static bool IsEditedStyle(WStyle style)
  {
    int keyWordLength = State.KeyWord.Length;
    string firstLetters = style.decoded.Substring(0, keyWordLength);
    if (firstLetters == State.KeyWord) 
      return style.decoded.Contains("+");
    return false;
  }
  
  public static void ParagraphCheck(Paragraph paragraph, int idx, ContentType type, WTable? table = null)
  {
    bool isParaExist = paragraph.ParagraphProperties != null;
    bool isInnerText = !string.IsNullOrEmpty(paragraph.InnerText);
    
    void onComment (string styleName, bool isEdited)
    {
      WReport.OnMessage(paragraph, type, idx, styleName, isEdited, table);
    }

    if (isInnerText)
    {
      if (isParaExist && paragraph.ParagraphProperties.ParagraphStyleId != null)
      {
        string styleName = paragraph.ParagraphProperties.ParagraphStyleId.Val;
        WStyle style = WStyle.GetStyleFromEncoded(styleName);
        
        // if the value of the pStyle is allowed => skip the paragraph
        if (style != null)
        {
          if (!IsValidStyle(style)) onComment(style.decoded, false);
          else if (IsEditedStyle(style)) onComment(style.decoded, true);
        }
        else onComment($"Undefined Style name '{styleName}'", false);
      }
      else if (isInnerText)onComment("Normal", false);
    }
  }
}