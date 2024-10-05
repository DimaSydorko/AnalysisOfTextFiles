using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AnalysisOfTextFiles.Objects;

public class CheckPage
{
  private static int SmInPoints(double sm)
  {
    return Convert.ToInt32(Math.Ceiling(sm * 567));
  }

  private static double PointsInSm(int points)
  {
    return Math.Round((double)points / 567, 2, MidpointRounding.ToEven);
  }

  private static readonly Dictionary<string, (int Width, int Height)> PageSizes = new Dictionary<string, (int, int)>
  {
    { "A3", (16838, 23811) },
    { "A4", (11906, 16838) },
    { "A5", (8391, 11906) },
    { "letter", (12240, 15840) }
  };

  public static void CheckDimensions(SectionProperties section)
  {
    var pageSize = section.GetFirstChild<PageSize>();
    if (pageSize == null) return;

    List<string> allowedSizes = State.PageSettings.Size;
    List<string> allowedOrientations = State.PageSettings.Orientation;

    // Check if the page size is valid according to the allowed sizes
    bool isValidSize = false;
    foreach (var size in allowedSizes)
    {
      if (PageSizes.ContainsKey(size))
      {
        var (allowedWidth, allowedHeight) = PageSizes[size];
        int pageWidth = Convert.ToInt32(pageSize.Width.Value);
        int pageHeight = Convert.ToInt32(pageSize.Height.Value);

        if ((pageWidth == allowedWidth && pageHeight == allowedHeight) ||
            (pageWidth == allowedHeight && pageHeight == allowedWidth))
        {
          isValidSize = true;
          break;
        }
      }
    }

    if (!isValidSize)
    {
      WReport.Write("Invalid page size. Allowed sizes are: " + string.Join(", ", allowedSizes));
    }

    // Check orientation
    var isLandscape = pageSize.Orient?.Value == PageOrientationValues.Landscape;
    var currentOrientation = isLandscape ? "landscape" : "portrait"; // 'h' for landscape, 'v' for portrait

    if (!allowedOrientations.Contains(currentOrientation))
    {
      WReport.Write(
        $"Invalid page orientation: '{currentOrientation}'. Allowed orientations are: {string.Join(", ", allowedOrientations)}");
    }
  }

  public static void CheckPageMargin(SectionProperties section)
  {
    var pageMargin = section.GetFirstChild<PageMargin>();
    WPage page = State.PageSettings;

    if (pageMargin == null) return;

    void CompareAndReport(string label, int actual, int expected)
    {
      if (actual != expected)
      {
        WReport.Write($"{label}: {PointsInSm(actual)} sm -> {PointsInSm(expected)} sm");
      }
    }

    CompareAndReport("Margin Top", pageMargin.Top, SmInPoints(page.MarginTop));
    CompareAndReport("Margin Bottom", pageMargin.Bottom, SmInPoints(page.MarginBottom));
    CompareAndReport("Margin Left", (int)pageMargin.Left.Value, SmInPoints(page.MarginLeft));
    CompareAndReport("Margin Right", (int)pageMargin.Right.Value, SmInPoints(page.MarginRight));

    CompareAndReport("Margin from Header", (int)pageMargin.Header.Value, SmInPoints(page.MarginHeader));
    CompareAndReport("Margin from Footer", (int)pageMargin.Footer.Value, SmInPoints(page.MarginFooter));
  }
}