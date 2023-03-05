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
    string first4Letters = style.decoded.Substring(0, 4);
    return allowedStyles.Contains(style.encoded) || first4Letters == "ЕОМ:";
  }

  public static void ParagraphCheck(Paragraph paragraph, int idx, ContentType type)
  {
    bool isParaExist = paragraph.ParagraphProperties != null;
    bool isInnerText = !string.IsNullOrEmpty(paragraph.InnerText);
    
    void onComment (string styleName)
    {
      WReport.OnMessage(paragraph, type, idx, styleName);
    }

    if (isInnerText)
    {
      // string first letters
      if (isParaExist && paragraph.ParagraphProperties.ParagraphStyleId != null)
      {
        string styleName = paragraph.ParagraphProperties.ParagraphStyleId.Val;
        WStyle style = WStyle.GetStyleFromEncoded(State.Styles, styleName);
        
        // if the value of the pStyle is allowed => skip the paragraph
        if (style != null)
        {
          if (!IsValidStyle(style)) onComment(style.decoded);
        }
        else onComment(styleName);
      }
      else if (isInnerText)onComment("Normal");
    }
  }
}