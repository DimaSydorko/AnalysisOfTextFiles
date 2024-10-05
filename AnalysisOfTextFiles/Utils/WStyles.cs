using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AnalysisOfTextFiles.Objects;

public class WStyles
{
  public static List<WStyle> GetDocStyles()
  {
    var styles = new List<WStyle>();
    var stylesXml = State.WDocument.MainDocumentPart.StyleDefinitionsPart.Styles;

    var data = AdminSettings.GetStyleData();
    State.Content = AdminSettings.GetStyleSettings(data);
    State.KeyWord = AdminSettings.GetStyleKeyWord(data);

    var list = stylesXml.ChildElements[1];

    //Map to get Encoded and Decoded StyleNames
    foreach (var styleXml in stylesXml.ChildElements)
      if (styleXml.ChildElements.Count >= 3)
      {
        var styleDec = styleXml.ChildElements[0];
        var styleEnc = styleXml.ChildElements[2];

        var style = new WStyle();

        void GetName(PropertyInfo property, OpenXmlElement styleXml, bool isDec)
        {
          if (property != null)
          {
            var styleNameObj = property.GetValue(styleXml);
            if (styleNameObj != null)
            {
              var propertyName = styleNameObj.GetType().GetProperty("Value");
              if (propertyName != null)
              {
                var name = propertyName.GetValue(styleNameObj)?.ToString();
                if (!string.IsNullOrEmpty(name))
                {
                  if (isDec) style.SetDec(WDecoding.RemoveSuffixIfExists(name));
                  else style.SetEnc(name);
                }
              }
            }
          }
        }

        var propertyDec = styleDec.GetType().GetProperty("Val");
        GetName(propertyDec, styleDec, true);

        var propertyEnc = styleEnc.GetType().GetProperty("Val");
        GetName(propertyEnc, styleEnc, false);

        if (style.Encoded == null && style.Decoded != null)
        {
          var keyWordLength = State.KeyWord.Length;
          var firstLetters = style.Decoded.Substring(0, keyWordLength);

          if (firstLetters == State.KeyWord)
          {
            var att = styleXml?.GetAttributes();
            if (att != null)
            {
              var styleIdAttr = att.FirstOrDefault(a => a.LocalName == "styleId");
              if (styleIdAttr != null) style.Encoded = styleIdAttr.Value;
            }
          }
        }

        //Rewrite TOC style names
        string[] tocStyles = { "toc 1", "toc 2", "toc 3", "TOC Heading" };
        if (tocStyles.Contains(style.Decoded))
        {
          var newEncoded = style.Decoded.Replace(" ", "");
          var Upper = newEncoded.Substring(0, 3).ToUpper() + newEncoded.Substring(3);
          style.Encoded = Upper;
        }

        //Rewrite Header style names
        var header = "Heading";
        if (!string.IsNullOrEmpty(style.Decoded) && style.Decoded.Length >= header.Length)
        {
          var firstHLetters = style.Decoded.Substring(0, header.Length);
          if (firstHLetters == header && style.Decoded.Length > header.Length + 1)
          {
            var hLevel = style.Decoded.Substring(header.Length + 1, 1);
            style.Encoded = $"{header}{hLevel}";
          }
        }

        //Save only styles which exist   
        var isNotAllowedStyle = style.Encoded != null && style.Encoded != "CommentText";

        var alreadyCreatedStyle = styles.FirstOrDefault(s => s.Encoded == style.Encoded);
        var isAlreadyCreated = alreadyCreatedStyle != null;

        if (style.Decoded != null && isNotAllowedStyle && !isAlreadyCreated) styles.Add(style);
        else if (style.Decoded == "Normal" && isNotAllowedStyle && !isAlreadyCreated) styles.Add(style);
      }

    return styles;
  }

  public static void Review()
  {
    WReport.Write("________Styles Review________");

    var styleDefinitionsPart = State.WDocument.MainDocumentPart.StyleDefinitionsPart;
    State.StylesSettings = StyleProperties.GetSettingsList();
    
    if (styleDefinitionsPart != null)
    {
      var stylesCheck = styleDefinitionsPart.Styles;
      foreach (var style in stylesCheck.Elements<Style>())
      {
        var wStyle = WStyle.GetStyleFromEncoded(style.StyleId);
        if (wStyle != null)
          if (style.StyleRunProperties != null && CheckParagraph.IsValidStyle(wStyle.Decoded))
          {
            var runProperties = style.StyleRunProperties;
            if (runProperties != null)
            {
              var properties = new StyleProperties();
              properties.name = wStyle.Decoded;

              if (runProperties.FontSize != null)
              {
                var val = runProperties.FontSize.Val.Value;
                var halfVal = int.Parse(val) / 2;
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

              var paragraphProperties = style.StyleParagraphProperties;
              if (paragraphProperties != null)
              {
                if (paragraphProperties?.SpacingBetweenLines != null)
                {
                  var spacing = paragraphProperties?.SpacingBetweenLines;
                  if (spacing != null)
                  {
                    properties.lineSpacingAfter = spacing.After?.Value ?? "0";
                    properties.lineSpacingBefore = spacing.After?.Value ?? "0";
                    properties.lineSpacing = spacing.Line?.Value ?? "0";
                  }
                }

                if (paragraphProperties?.TextAlignment != null)
                  properties.position = paragraphProperties?.TextAlignment?.Val;
              }
              else
              {
                properties.lineSpacingAfter = "0";
                properties.lineSpacingBefore = "0";
                properties.lineSpacing = "1.5";
              }

              if (runProperties.RunFonts != null && runProperties.RunFonts.Ascii != null)
                properties.fontType = runProperties.RunFonts.Ascii.InnerText;
              else
                properties.fontType = "Times New Roman";

              var settings = State.StylesSettings.FirstOrDefault(s => s.name == properties.name);

              if (settings != null)
              {
                var diff = WReport.OnCompareStyleSettings(settings, properties);
                if (!string.IsNullOrEmpty(diff))
                {
                  WReport.Write($"\n[{wStyle.Decoded}]");
                  WReport.Write(diff);
                }
              }
            }
          }
      }
    }

    WReport.Write( "________Content Review________");
  }
}