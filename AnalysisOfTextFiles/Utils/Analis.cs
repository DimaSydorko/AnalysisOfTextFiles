using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AnalysisOfTextFiles.Objects;

public class Analis
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

  public static bool IsValidStyle(WStyle style)
  {
    var keyWordLength = State.KeyWord.Length;
    if (keyWordLength <= style.decoded.Length)
    {
      var firstLetters = style.decoded.Substring(0, keyWordLength);
      return allowedStyles.Contains(style.encoded) || firstLetters == State.KeyWord;
    }

    return false;
  }

  public static string? GetOldDecStyle(string encoded)
  {
    var entryTable = new Dictionary<string, string>
    {
      { "21", "TOC1" },
      { "22", "TOC2" },
      { "23", "TOC3" },
      { "13", "Normal" },
      { "1", "Heading1" },
      { "2", "Heading2" },
      { "3", "Heading3" }
    };

    if (entryTable.ContainsKey(encoded))
    {
      var entry = entryTable[encoded];
      return entry;
    }

    return null;
  }

  public static bool IsEditedStyle(WStyle style)
  {
    var keyWordLength = State.KeyWord.Length;
    var firstLetters = style.decoded.Substring(0, keyWordLength);
    if (firstLetters == State.KeyWord)
      return style.decoded.Contains("+");
    return false;
  }

  public static void ParagraphCheck(Paragraph paragraph, int idx, ContentType type, WTable? table = null)
  {
    var isParaExist = paragraph.ParagraphProperties != null;
    var isInnerText = !string.IsNullOrEmpty(paragraph.InnerText);

    void onComment(string styleName, bool isEdited)
    {
      WReport.OnMessage(paragraph, type, idx, styleName, isEdited, table);
    }

    if (isInnerText)
    {
      if (isParaExist && paragraph.ParagraphProperties.ParagraphStyleId != null)
      {
        string styleName = paragraph.ParagraphProperties.ParagraphStyleId.Val;
        var style = WStyle.GetStyleFromEncoded(styleName);

        // if the value of the pStyle is allowed => skip the paragraph
        if (style != null)
        {
          if (!IsValidStyle(style)) onComment(style.decoded, false);
          else if (IsEditedStyle(style)) onComment(style.decoded, true);
        }
        else
        {
          var dec = GetOldDecStyle(styleName);
          if (!allowedStyles.Contains(dec))
          {
            if (dec == null) onComment($"Undefined Style name '{styleName}'", false);
            else
              onComment(dec, false);
          }
        }
      }
      else if (isInnerText)
      {
        onComment("Normal", false);
      }
    }
    else
    {
      WReport.OnMessage(paragraph, type, idx, "", false, table, "Empty line");
    }
  }

  public static void CheckDimensions(SectionProperties section)
  {
    var pageSize = section.GetFirstChild<PageSize>();
    if (pageSize == null) return;
    
    var isLetter = pageSize.Width == 12240 && pageSize.Height == 15840;
    var isA4 = pageSize.Width == 11906 && pageSize.Height == 16838;
    
    if (!isLetter && !isA4)
    {
      WReport.Write("Invalid page size, should be 'letter' or 'A4'");
    }

    var isLandscape = pageSize.Orient?.Value == PageOrientationValues.Landscape;
    
    if (isLandscape)
    {
      WReport.Write("Invalid page orientation: 'Landscape'");
    }
  }

  public static void CheckPageMargin(SectionProperties section)
  {
    var pageMargin = section.GetFirstChild<PageMargin>();
    const int mTop = 1418, mBottom = 851, mLeft = 1134, mRight = 1134, mFooter = 709, mHeader = 709;
    
    if (pageMargin == null) return;
    
    double PointsIntoSm(int points) => Math.Round((double)points / 567, 2, MidpointRounding.ToEven);
    void CompareAndReport(string label, int actual, int expected)
    {
      if (actual != expected)
      {
        WReport.Write($"{label}: {PointsIntoSm(actual)} sm -> {PointsIntoSm(expected)} sm");
      }
    }

    CompareAndReport("Margins Top", pageMargin.Top, mTop);
    CompareAndReport("Margins Bottom", pageMargin.Bottom, mBottom);
    CompareAndReport("Margins Left", (int)pageMargin.Left.Value, mLeft);
    CompareAndReport("Margins Right", (int)pageMargin.Right.Value, mRight);

    CompareAndReport("Margin from Header", (int)pageMargin.Header.Value, mHeader);
    CompareAndReport("Margin from Footer", (int)pageMargin.Footer.Value, mFooter);
  }
}