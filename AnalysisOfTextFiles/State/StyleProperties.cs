using System;
using System.Collections.Generic;
using System.Linq;

public class StyleProperties
{
  public string name { get; set; }
  public string size { get; set; }
  public string position { get; set; }
  public string lineSpacing { get; set; }
  public string lineSpacingBefore { get; set; }
  public string lineSpacingAfter { get; set; }

  public string color { get; set; }
  public string fontType { get; set; }

  public string bold { get; set; }
  public string italic { get; set; }
  public string underline { get; set; }
  public string capitalize { get; set; }

  public List<string> after { get; set; }

  public List<string> before { get; set; }


  public static List<StyleProperties> GetSettingsList()
  {
    var stylesSettings = new List<StyleProperties>();

    var lines = State.Content.Split(new[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);

    List<string> getOrder(string value)
    {
      return value
        .Split(',')
        .Select(s => s.Trim().Trim('"'))
        .ToList();
    }

    if (lines.Length != 0)
    {
      StyleProperties currentStyle = null;
      foreach (var line in lines)
        if (line.StartsWith("["))
        {
          // Start of a new section, create a new Style object
          currentStyle = new StyleProperties();
          currentStyle.name = line.Substring(1, line.Length - 2);
          stylesSettings.Add(currentStyle);
        }
        else if (currentStyle != null)
        {
          // Parse the key-value pairs and set the properties of the current Style object
          var parts = line.Split('=');
          if (parts.Length == 2)
          {
            var key = parts[0].Trim();
            var value = parts[1].Trim();

            switch (key)
            {
              case "size":
                currentStyle.size = value;
                break;
              case "position":
                currentStyle.position = value;
                break;
              case "lineSpacing":
                currentStyle.lineSpacing = value;
                break;
              case "lineSpacingBefore":
                currentStyle.lineSpacingBefore = value;
                break;
              case "lineSpacingAfter":
                currentStyle.lineSpacingAfter = value;
                break;
              case "color":
                currentStyle.color = value;
                break;
              case "fontType":
                currentStyle.fontType = value;
                break;
              case "bold":
                currentStyle.bold = value;
                break;
              case "italic":
                currentStyle.italic = value;
                break;
              case "underline":
                currentStyle.underline = value;
                break;
              case "capitalize":
                currentStyle.capitalize = value;
                break;
              case "before":
                // var before = getOrder(value);
                // before.Add(currentStyle.name);
                currentStyle.before = getOrder(value);
                break;
              case "after":
                // var after = getOrder(value);
                // after.Add(currentStyle.name);
                currentStyle.after = getOrder(value);
                break;
            }
          }
        }
    }

    return stylesSettings;
  }
}