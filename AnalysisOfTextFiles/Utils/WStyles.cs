using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Reflection;
using DocumentFormat.OpenXml;

namespace AnalysisOfTextFiles.Objects;

public class WStyles
{
    public static List<WStyle> GetDocStyles()
  {
    List<WStyle> styles = new List<WStyle>();
    Styles stylesXml = State.WDocument.MainDocumentPart.StyleDefinitionsPart.Styles;
    
    //Map to get Encoded and Decoded StyleNames
    foreach (var styleXml in stylesXml.ChildElements)
    {
      if (styleXml.ChildElements.Count >= 3)
      {
        OpenXmlElement styleDec = styleXml.ChildElements[0];
        OpenXmlElement styleEnc = styleXml.ChildElements[2];

        WStyle style = new WStyle();

        void GetName (PropertyInfo property, OpenXmlElement styleXml, bool isDec)
        {
          if (property != null)
          {
            var styleNameObj = property.GetValue(styleXml);
            if (styleNameObj != null)
            {
              PropertyInfo propertyName = styleNameObj.GetType().GetProperty("Value");
              if (propertyName != null)
              {
                string name = propertyName.GetValue(styleNameObj)?.ToString();
                if (!string.IsNullOrEmpty(name))
                {
                  if (isDec) style.SetDec(name);
                  else style.SetEnc(name);
                }
              }
            }
          }
        }
        
        PropertyInfo propertyDec = styleDec.GetType().GetProperty("Val");
        GetName(propertyDec, styleDec, true);

        PropertyInfo propertyEnc = styleEnc.GetType().GetProperty("Val");
        GetName(propertyEnc, styleEnc, false);

        //Rewrite TOC style names
        string[] tocStyles = { "toc 1", "toc 2", "toc 3", "TOC Heading" };
        if (tocStyles.Contains(style.decoded))
        {
          string newEncoded = style.decoded.Replace(" ", "");
          string Upper = newEncoded.Substring(0, 3).ToUpper() + newEncoded.Substring(3);
          style.encoded = Upper;
        }

        //Save only styles which exist   
        bool isNotAllowedStyle = style.encoded != null && style.encoded != "CommentText";

        WStyle alreadyCreatedStyle = styles.FirstOrDefault(s => s.encoded == style.encoded);
        bool isAlreadyCreated = alreadyCreatedStyle != null; 
          
        if (style.decoded != null && isNotAllowedStyle && !isAlreadyCreated) styles.Add(style);
        else if (style.decoded == "Normal" && isNotAllowedStyle && !isAlreadyCreated) styles.Add(style);
      }
    }
    
    return styles;
  }
}