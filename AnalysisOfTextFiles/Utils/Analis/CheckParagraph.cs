using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
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
  
  public static List<string> allowedStyles = new() { "Heading1", "Heading2", "Heading3", "TOC1", "TOC2" };

  public static bool IsValidWStyle(WStyle style)
  {
    bool isInSettings = Convert.ToBoolean(State.StylesSettings.Exists(s => s.name == style.Decoded));

    if (isInSettings)
    {
      return true;
    }
    
    var keyWordLength = State.KeyWord.Length;
    if (keyWordLength <= style.Decoded.Length)
    {
      var firstLetters = style.Decoded.Substring(0, keyWordLength);

      return allowedStyles.Contains(style.Encoded) || firstLetters == State.KeyWord;
    }

    return false;
  }

  public static string GetParagraphStyle(Paragraph? paragraph)
  {
    if (paragraph == null) return "Unknown";

    bool isParExist = paragraph.ParagraphProperties != null && paragraph.ParagraphProperties?.ParagraphStyleId?.Val != null;

    if (isParExist)
    {
      string? styleName = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? paragraph.ParagraphProperties?.ParagraphStyleId?.Val;
      var style = WStyle.GetDecodedStyle(styleName);
      return style;
    }

    return "Normal";
  }
  
  public static WStyle GetParagraphWStyle(Paragraph? paragraph)
  {
    string styleName = paragraph.ParagraphProperties.ParagraphStyleId.Val.Value ?? paragraph.ParagraphProperties.ParagraphStyleId.Val;
    WStyle style = WStyle.GetStyleFromEncoded(styleName);
    
    return style;
  }

  public static void ParagraphCheck(Paragraph paragraph, Paragraph prevParagraph, Paragraph nextParagraph, int idx, ContentType type, WTable? table = null)
  {
    var isParaExist = paragraph.ParagraphProperties != null;
    var hasInnerText = !string.IsNullOrEmpty(paragraph.InnerText);
    var hasImage = paragraph.Descendants<Drawing>().Any() || paragraph.Descendants<Inline>().Any();

    void onComment(string styleName, WReport.TitleType? titleType = WReport.TitleType.Wrong)
    {
      WReport.OnMessage(paragraph, type, idx, styleName, table, titleType);
    }

    if (hasInnerText || hasImage)
    {
      if (isParaExist && paragraph.ParagraphProperties?.ParagraphStyleId != null)
      {
        string? styleName = WDecoding.RemoveSuffixIfExists(GetParagraphStyle(paragraph));
        Order.CheckParagraph(paragraph, prevParagraph, nextParagraph, type, styleName, idx);

        WStyle style = GetParagraphWStyle(paragraph);
        // if the value of the pStyle is allowed => skip the paragraph
        if (style != null)
        {
          if (!IsValidWStyle(style)) onComment(style.Decoded);
          else if (CheckEdited.IsEditedStyle(paragraph)) onComment(style.Decoded, WReport.TitleType.Edited);
        }
        else
        {
          string dec = WDecoding.GetOldDecStyle(styleName);
          if (!allowedStyles.Contains(dec))
          {
            if (dec == null) onComment($"Undefined Style name '{styleName}'");
            else onComment(dec);
          }
        }
      }
      else
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