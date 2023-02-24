using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AnalysisOfTextFiles.Objects;

class Analis
{
  public static bool IsValidStyle(WStyle style)
  {
    string[] allowedStyles = { "Heading1", "Heading2", "Heading3" };

    string first4Letters = style.decoded.Substring(0, 4);
    return allowedStyles.Contains(style.encoded) || first4Letters == "ЕОМ:";
  }

  public static void ParagraphCheck(Paragraph paragraph, MainDocumentPart mainPart,  List<WStyle> docStyles)
  {
    bool isParaExist = paragraph.ParagraphProperties != null;
    bool isInnerText = !string.IsNullOrEmpty(paragraph.InnerText);
    
    if (isParaExist || isInnerText)
    {
      var styleId = paragraph.ParagraphProperties.ParagraphStyleId;
      if (isParaExist && styleId != null)
      {
        string styleName = styleId.Val;
        WStyle style = WStyle.GetStyleFromEncoded(docStyles, styleName);

        // if the value of the pStyle is allowed => skip the paragraph
        if (style != null && !IsValidStyle(style))
        {
          WComment.Add(mainPart, paragraph, style.decoded);
        }
      }
      else if (isInnerText)
      {
        WComment.Add(mainPart, paragraph, "Normal");
      }
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

        bool isNotAllowedStyle = style.encoded != null && style.encoded != "CommentText";
        
        if (style.decoded != null && isNotAllowedStyle)
        {
          styles.Add(style);
        } else if (style.decoded == "Normal" && isNotAllowedStyle)
        {
          styles.Add(style);
        }
      }
    }

    return styles;
  }
}