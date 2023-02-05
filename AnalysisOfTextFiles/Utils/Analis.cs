using System.Reflection;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AnalysisOfTextFiles.Objects;

class Analis
{
  public static List<WStyle> ExtractStyles(Styles stylesXml)
  {
    List<WStyle> styles = new List<WStyle>();
    // WStyle[] styles = Array.Empty<WStyle>();

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
            string last5Letter = decName.Substring(decName.Length - 5, 5);
            if (last5Letter == " Char")
            {
              string withoutLast5Letter = decName.Substring(0, decName.Length - 5);
              style.SetDec(withoutLast5Letter);
            }
            else style.SetDec(decName);
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

        if (style.decoded != null)
        {
          styles.Add(style);
        }
      }
    }

    return styles;
  }
}