using System.Reflection;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System.Linq;

namespace AnalysisOfTextFiles.Objects;

class Analis
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

  public static void ParagraphCheck(Paragraph paragraph, MainDocumentPart mainPart, List<WStyle> docStyles,
    string reportPath, int idx, ContentType type)
  {
    string text = paragraph.InnerText;
    bool isParaExist = paragraph.ParagraphProperties != null;
    bool isInnerText = !string.IsNullOrEmpty(text);
    bool isComment = type != ContentType.Header || type != ContentType.Footer;

    void onComment (string styleName, string first10Letters)
    {
      if (isComment) WComment.Add(mainPart, paragraph, styleName);
      string typeTitle = type == ContentType.TOC ? "Table of content" : $"{type} №{idx + 1}";
      File.AppendAllText(reportPath, $"{typeTitle} ('{first10Letters}') Style: {styleName}\n");
    }

    if (isParaExist || isInnerText)
    {
      // string first 10 letter
      string first10Letters = text.Length > 10 ? text.Substring(0, 10) + "..." : text;
      if (isParaExist && paragraph.ParagraphProperties.ParagraphStyleId != null)
      {
        var styleId = paragraph.ParagraphProperties.ParagraphStyleId;

        string styleName = styleId.Val;
        WStyle style = WStyle.GetStyleFromEncoded(docStyles, styleName);

        // if the value of the pStyle is allowed => skip the paragraph
        if (style != null && !IsValidStyle(style)) onComment(style.decoded, first10Letters);
      }
      else if (isInnerText)onComment("Normal", first10Letters);
    }
  }

  public static List<WStyle> ExtractStyles(Styles stylesXml)
  {
    List<WStyle> styles = new List<WStyle>();

    //Map to get Encoded and Decoded StyleNames
    foreach (var styleXml in stylesXml.ChildElements)
    {
      if (styleXml.ChildElements.Count >= 3)
      {
        var styleDec = styleXml.ChildElements[0];
        var styleEnc = styleXml.ChildElements[2];

        WStyle style = new WStyle();

        //Get Decoded Name
        PropertyInfo propertyDec = styleDec.GetType().GetProperty("Val");
        if (propertyDec != null)
        {
          var styleNameObjDec = propertyDec.GetValue(styleDec);
          PropertyInfo propertyDecName = styleNameObjDec.GetType().GetProperty("Value");
          string decName = propertyDecName.GetValue(styleNameObjDec).ToString();

          //Remove " Char" from styleName 
          if (decName.Length > 5)
          {
            // string last5Letter = decName.Substring(decName.Length - 5, 5);
            // if (last5Letter == " Char")
            // {
            //   string withoutLast5Letter = decName.Substring(0, decName.Length - 5);
            //   style.SetDec(withoutLast5Letter);
            // }
            // else 
            style.SetDec(decName);
          }
          else style.SetDec(decName);
        }

        //Get Encoded Name
        PropertyInfo propertyEnc = styleEnc.GetType().GetProperty("Val");
        if (propertyEnc != null)
        {
          var styleNameObjEnc = propertyEnc.GetValue(styleEnc);
          if (styleNameObjEnc != null)
          {
            PropertyInfo propertyEncName = styleNameObjEnc.GetType().GetProperty("Value");
            string encName = propertyEncName.GetValue(styleNameObjEnc).ToString();

            style.SetEnc(encName);
          }
        }

        string[] tocStyles = { "toc 1", "toc 2", "toc 3", "TOC Heading" };
        if (tocStyles.Contains(style.decoded) )
        {
          string newEncoded = style.decoded.Replace(" ", "");
          string Upper = newEncoded.Substring(0, 3).ToUpper() + newEncoded.Substring(3);
          style.encoded = Upper;
        }

        bool isNotAllowedStyle = style.encoded != null && style.encoded != "CommentText";

        if (style.decoded != null && isNotAllowedStyle)
        {
          styles.Add(style);
        }
        else if (style.decoded == "Normal" && isNotAllowedStyle)
        {
          styles.Add(style);
        }
      }
    }

    return styles;
  }
}