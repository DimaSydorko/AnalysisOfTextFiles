using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
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
    var isInSettings = Convert.ToBoolean(State.StylesSettings.Exists(s => s.name == style.Decoded));

    if (isInSettings) return true;

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

    var isParExist = paragraph.ParagraphProperties != null &&
                     paragraph.ParagraphProperties?.ParagraphStyleId?.Val != null;

    if (isParExist)
    {
      var styleName = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value ??
                      paragraph.ParagraphProperties?.ParagraphStyleId?.Val;
      var style = WStyle.GetDecodedStyle(styleName);
      return style;
    }

    return "Normal";
  }

  public static WStyle GetParagraphWStyle(Paragraph? paragraph)
  {
    var styleName = paragraph.ParagraphProperties.ParagraphStyleId.Val.Value ??
                    paragraph.ParagraphProperties.ParagraphStyleId.Val;
    var style = WStyle.GetStyleFromEncoded(styleName);

    return style;
  }

  private static void StyleCheck(Paragraph? paragraph, string styleName, int idx, ContentType type,
    WTable? table = null)
  {
    void onComment(string styleName, WReport.TitleType? titleType = WReport.TitleType.Wrong)
    {
      WReport.OnMessage(paragraph, type, idx, styleName, table, titleType);
    }

    var style = GetParagraphWStyle(paragraph);
    // if the value of the pStyle is allowed => skip the paragraph
    if (style != null)
    {
      if (!IsValidWStyle(style)) onComment(style.Decoded);
      else if (CheckEdited.IsEditedStyle(paragraph)) onComment(style.Decoded, WReport.TitleType.Edited);
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

  public static async Task ParagraphCheck(Paragraph paragraph, int idx, ContentType type, WTable? table = null)
  {
    var isParaExist = paragraph.ParagraphProperties != null;
    var hasInnerText = !string.IsNullOrEmpty(paragraph.InnerText);
    var hasImage = paragraph.Descendants<Drawing>().Any() || paragraph.Descendants<Inline>().Any();

    if (hasInnerText || hasImage)
    {
      if (isParaExist && paragraph.ParagraphProperties?.ParagraphStyleId != null)
      {
        var styleName = WDecoding.RemoveSuffixIfExists(GetParagraphStyle(paragraph));
        // await Task.Delay(1);

        Order.CheckParagraph(paragraph, type, styleName, idx, table);
        StyleCheck(paragraph, styleName, idx, type, table);
      }
      else
      {
        WReport.OnMessage(paragraph, type, idx, "Normal", table);
      }
    }
    else
    {
      if (!State.IsAllowEmptyLine)
      {
        WReport.OnMessage(paragraph, type, idx, "", table, WReport.TitleType.Empty);
      }
    }
  }
}