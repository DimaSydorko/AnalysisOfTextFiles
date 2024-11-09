using System;
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
    foreach (var styleXml in stylesXml.Elements<Style>())
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


        var id = styleXml?.StyleId?.Value;
        if (id != null)
        {
          style.SetEnc(id);
        }
        else
        {
          var propertyEnc = styleEnc.GetType().GetProperty("Val");
          GetName(propertyEnc, styleEnc, false);
        }

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

  private static string SpacingDecoding(string str)
  {
    return Convert.ToString(Convert.ToDouble(str) / 240);
  }

  private static StyleProperties _getStyleProperties(Style style)
  {
    var styleVal = style?.StyleId?.Value;
    var wStyle = WStyle.GetStyleFromEncoded(styleVal ?? style?.StyleId);
    var runProperties = style.StyleRunProperties;
    var properties = new StyleProperties();

    properties.name = wStyle.Decoded;

    if (runProperties.FontSize != null)
    {
      var val = runProperties.FontSize.Val.Value;
      var halfVal = int.Parse(val) / 2;
      properties.size = $"{halfVal}";
    }

    properties.color = runProperties.Color?.Val ?? "000000";
    properties.bold = runProperties.Bold != null ? "true" : "false";
    properties.italic = runProperties.Italic != null ? "true" : "false";
    properties.underline = runProperties.Underline != null ? "true" : "false";
    properties.capitalize = runProperties.Caps != null ? "true" : "false";

    properties.position = runProperties.Position?.Val ?? "Left";

    var paragraphProperties = style.StyleParagraphProperties;
    if (paragraphProperties != null)
    {
      if (paragraphProperties.SpacingBetweenLines != null)
      {
        var spacing = paragraphProperties.SpacingBetweenLines;
        properties.lineSpacingAfter = SpacingDecoding(spacing.After?.Value ?? "0");
        properties.lineSpacingBefore = SpacingDecoding(spacing.Before?.Value ?? "0");
        properties.lineSpacing = SpacingDecoding(spacing.Line?.Value ?? "0");
      }

      properties.position = paragraphProperties.Justification?.Val?.Value.ToString() ?? properties.position;
    }

    properties.lineSpacingAfter = properties.lineSpacingAfter ?? "0";
    properties.lineSpacingBefore = properties.lineSpacingBefore ?? "0";
    properties.lineSpacing = properties.lineSpacing ?? "1.5";
    properties.fontType = runProperties.RunFonts?.Ascii?.InnerText ?? "Times New Roman";

    return properties;
  }

  public static void Review()
  {
    WReport.Write("________Styles Review________");

    var styleDefinitionsPart = State.WDocument.MainDocumentPart.StyleDefinitionsPart;
    State.StylesSettings = StyleProperties.GetSettingsList();
    State.PageSettings = PageProperties.GetPageSettings();
    List<string> alreadyChecked = new List<string>();

    if (styleDefinitionsPart != null)
    {
      var stylesCheck = styleDefinitionsPart.Styles;
      foreach (var style in stylesCheck.Elements<Style>())
      {
        var styleVal = style?.StyleId?.Value;
        var wStyle = WStyle.GetStyleFromEncoded(styleVal ?? style?.StyleId);

        if (wStyle != null)
        {
          bool isChecked = alreadyChecked.Contains(wStyle.Decoded);
          if (!isChecked && style.StyleRunProperties != null && CheckParagraph.IsValidWStyle(wStyle))
          {
            var properties = _getStyleProperties(style);
            var settings = State.StylesSettings.FirstOrDefault(s => s.name == properties.name);
            alreadyChecked.Add(wStyle.Decoded);

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

    WReport.Write("________Content Review________");
  }

  public static void AnaliseStylesSettings()
  {
    var styleDefinitionsPart = State.WDocument.MainDocumentPart.StyleDefinitionsPart;
    if (styleDefinitionsPart != null)
    {
      var stylesCheck = styleDefinitionsPart.Styles;
      List<string> alreadyChecked = new List<string>();
      foreach (var style in stylesCheck.Elements<Style>())
      {
        var styleVal = style?.StyleId?.Value;
        var wStyle = WStyle.GetStyleFromEncoded(styleVal ?? style?.StyleId);
        if (wStyle != null)
        {
          bool isChecked = alreadyChecked.Contains(wStyle.Decoded);

          if (!isChecked && style.StyleRunProperties != null)
          {
            var properties = _getStyleProperties(style);
            alreadyChecked.Add(wStyle.Decoded);

            void Write(string message)
            {
              WReport.Write($"{message}", false, true);
            }

            Write($"[{properties.name}]");
            Write($"name={properties.name}");
            Write($"size={properties.size}");
            Write($"position={properties.position}");
            Write($"lineSpacing={properties.lineSpacing}");
            Write($"lineSpacingBefore={properties.lineSpacingBefore}");
            Write($"lineSpacingAfter={properties.lineSpacingAfter}");
            Write($"color={properties.color}");
            Write($"fontType={properties.fontType}");
            Write($"bold={properties.bold}");
            Write($"italic={properties.italic}");
            Write($"underline={properties.underline}");
            Write($"capitalize={properties.capitalize}");
            Write("\n");
          }
        }
      }
    }
  }
}