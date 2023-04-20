using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

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

        void GetName(PropertyInfo property, OpenXmlElement styleXml, bool isDec)
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

  public static void Review()
  {
    File.AppendAllText(State.FilePath.report, "________Styles Review________\n");

    StyleDefinitionsPart? styleDefinitionsPart = State.WDocument.MainDocumentPart.StyleDefinitionsPart;
    List<StyleProperties> stylesSettings = StyleProperties.GetSettingsList();

    if (styleDefinitionsPart != null)
    {
      Styles stylesCheck = styleDefinitionsPart.Styles;
      foreach (Style style in stylesCheck.Elements<Style>())
      {
        WStyle wStyle = WStyle.GetStyleFromEncoded(style.StyleId);
        if (wStyle != null)
        {
          if (style.StyleRunProperties != null && Analis.IsValidStyle(wStyle))
          {
            StyleRunProperties? runProperties = style.StyleRunProperties;
            if (runProperties != null)
            {
              StyleProperties properties = new StyleProperties();
              properties.name = wStyle.decoded;

              if (runProperties.FontSize != null)
              {
                string val = runProperties.FontSize.Val.Value;
                int halfVal = int.Parse(val) / 2;
                properties.size = $"{halfVal}";
              }

              if (runProperties.Color != null)
                properties.color = runProperties.Color?.Val ?? "000000";
              else
                properties.color = "000000";

              if (runProperties.Position != null)
                properties.position = runProperties.Position.Val;
              else
                properties.position = "center";

              if (runProperties.Bold != null)
                properties.bold = "true";
              else
                properties.bold = "false";

              if (runProperties.Italic != null)
                properties.italic = "true";
              else
                properties.italic = "false";

              if (runProperties.Underline != null)
                properties.underline = "true";
              else
                properties.underline = "false";

              if (runProperties.Caps != null)
                properties.capitalize = "true";
              else
                properties.capitalize = "false";

              StyleParagraphProperties? paragraphProperties = style.StyleParagraphProperties;
              if (paragraphProperties != null)
              {
                if (paragraphProperties?.SpacingBetweenLines != null)
                {
                  SpacingBetweenLines? spacing = paragraphProperties?.SpacingBetweenLines;
                  if (spacing != null)
                  {
                    properties.lineSpacingAfter = spacing.After?.Value ?? "0";
                    properties.lineSpacingBefore = spacing.After?.Value ?? "0";
                    properties.lineSpacing = spacing.Line?.Value ?? "0";
                  }
                }

                if (paragraphProperties?.TextAlignment != null)
                {
                  properties.position = paragraphProperties?.TextAlignment?.Val;
                }
              }
              else
              {
                properties.lineSpacingAfter = "0";
                properties.lineSpacingBefore = "0";
                properties.lineSpacing = "1.5";
              }

              if (runProperties.RunFonts != null && runProperties.RunFonts.Ascii != null)
                properties.fontType = runProperties.RunFonts.Ascii.InnerText;

              var settings = stylesSettings.FirstOrDefault(s => s.name == properties.name);

              if (settings != null)
              {
                string diff = WReport.OnCompareObjects(settings, properties);
                if (!string.IsNullOrEmpty(diff))
                {
                  File.AppendAllText(State.FilePath.report, $"\n[{wStyle.decoded}]\n");
                  File.AppendAllText(State.FilePath.report, diff);
                }
              }
            }
          }
        }
      }
    }

    File.AppendAllText(State.FilePath.report, "________Content Review________\n");
  }
}