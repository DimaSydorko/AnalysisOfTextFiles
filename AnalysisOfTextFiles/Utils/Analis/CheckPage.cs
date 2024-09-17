using System;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AnalysisOfTextFiles.Objects;

public class CheckPage
{
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