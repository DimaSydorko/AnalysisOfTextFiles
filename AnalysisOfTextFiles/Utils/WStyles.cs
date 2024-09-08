using System.Collections.Generic;
using System.IO;
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
                  if (isDec) style.SetDec(name);
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

        if (style.encoded == null && style.decoded != null)
        {
          var keyWordLength = State.KeyWord.Length;
          var firstLetters = style.decoded.Substring(0, keyWordLength);

          if (firstLetters == State.KeyWord)
          {
            var att = styleXml?.GetAttributes();
            if (att != null)
            {
              var styleIdAttr = att.FirstOrDefault(a => a.LocalName == "styleId");
              if (styleIdAttr != null) style.encoded = styleIdAttr.Value;
            }
          }
        }

        //Rewrite TOC style names
        string[] tocStyles = { "toc 1", "toc 2", "toc 3", "TOC Heading" };
        if (tocStyles.Contains(style.decoded))
        {
          var newEncoded = style.decoded.Replace(" ", "");
          var Upper = newEncoded.Substring(0, 3).ToUpper() + newEncoded.Substring(3);
          style.encoded = Upper;
        }

        //Rewrite Header style names
        var header = "Heading";
        if (!string.IsNullOrEmpty(style.decoded) && style.decoded.Length >= header.Length)
        {
          var firstHLetters = style.decoded.Substring(0, header.Length);
          if (firstHLetters == header && style.decoded.Length > header.Length + 1)
          {
            var hLevel = style.decoded.Substring(header.Length + 1, 1);
            style.encoded = $"{header}{hLevel}";
          }
        }

        //Save only styles which exist   
        var isNotAllowedStyle = style.encoded != null && style.encoded != "CommentText";

        var alreadyCreatedStyle = styles.FirstOrDefault(s => s.encoded == style.encoded);
        var isAlreadyCreated = alreadyCreatedStyle != null;

        if (style.decoded != null && isNotAllowedStyle && !isAlreadyCreated) styles.Add(style);
        else if (style.decoded == "Normal" && isNotAllowedStyle && !isAlreadyCreated) styles.Add(style);
      }

    return styles;
  }

  public static void Review()
  {
    File.AppendAllText(State.FilePath.report, "________Styles Review________\n");

    var styleDefinitionsPart = State.WDocument.MainDocumentPart.StyleDefinitionsPart;
    var stylesSettings = StyleProperties.GetSettingsList();

    if (styleDefinitionsPart != null)
    {
      var stylesCheck = styleDefinitionsPart.Styles;
      foreach (var style in stylesCheck.Elements<Style>())
      {
        var wStyle = WStyle.GetStyleFromEncoded(style.StyleId);
        if (wStyle != null)
          if (style.StyleRunProperties != null && Analis.IsValidStyle(wStyle))
          {
            var runProperties = style.StyleRunProperties;
            if (runProperties != null)
            {
              var properties = new StyleProperties();
              properties.name = wStyle.decoded;

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

              var settings = stylesSettings.FirstOrDefault(s => s.name == properties.name);

              if (settings != null)
              {
                var diff = WReport.OnCompareObjects(settings, properties);
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

    File.AppendAllText(State.FilePath.report, "________Content Review________\n");
  }
}