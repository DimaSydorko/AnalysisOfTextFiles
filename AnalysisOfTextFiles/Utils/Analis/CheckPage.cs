using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AnalysisOfTextFiles.Objects;

public class CheckPage
{
  private static int CmInPoints(double cm)
  {
    return Convert.ToInt32(Math.Ceiling(cm * 567));
  }

  private static double PointsInCm(int points)
  {
    return Math.Round((double)points / 567, 2, MidpointRounding.ToEven);
  }

  public static bool IsEqual(int num1, int num2, int tolerance = 5)
  {
    return Math.Abs(num1 - num2) <= tolerance;
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

        if ((IsEqual(pageWidth, allowedWidth) && IsEqual(pageHeight, allowedHeight)) ||
            (IsEqual(pageWidth, allowedHeight) && IsEqual(pageHeight, allowedWidth)))
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
      if (!IsEqual(actual, expected, 3))
      {
        WReport.Write($"{label}: {PointsInCm(actual)} cm -> {PointsInCm(expected)} cm");
      }
    }

    CompareAndReport("Margin Top", pageMargin.Top, CmInPoints(page.MarginTop));
    CompareAndReport("Margin Bottom", pageMargin.Bottom, CmInPoints(page.MarginBottom));
    CompareAndReport("Margin Left", (int)pageMargin.Left.Value, CmInPoints(page.MarginLeft));
    CompareAndReport("Margin Right", (int)pageMargin.Right.Value, CmInPoints(page.MarginRight));

    CompareAndReport("Margin from Header", (int)pageMargin.Header.Value, CmInPoints(page.MarginHeader));
    CompareAndReport("Margin from Footer", (int)pageMargin.Footer.Value, CmInPoints(page.MarginFooter));
  }

  public static void AnalisePageSettings()
  {
    WReport.WriteSettings("{PAGE}");

    var body = State.WDocument.MainDocumentPart.Document.Body;
    var section = body.GetFirstChild<SectionProperties>();

    var pageSize = section.GetFirstChild<PageSize>();
    if (pageSize != null)
    {
      int pageWidth = Convert.ToInt32(pageSize.Width.Value);
      int pageHeight = Convert.ToInt32(pageSize.Height.Value);

      var sizeEntry = PageSizes.FirstOrDefault(p =>
        (IsEqual(p.Value.Width, pageWidth) && IsEqual(p.Value.Height, pageHeight)) ||
        (IsEqual(p.Value.Height, pageWidth) && IsEqual(p.Value.Width, pageHeight))
      );

      string sizeName = sizeEntry.Key ?? "Custom";
      WReport.WriteSettings($"pageSize={sizeName}");

      string orientation = pageSize.Orient == null ? "portrait" :
        pageSize.Orient == PageOrientationValues.Landscape ? "landscape" : "portrait";
      WReport.WriteSettings($"orientation={orientation}");
    }

    var pageMargin = section.GetFirstChild<PageMargin>();
    if (pageMargin != null)
    {
      WReport.WriteSettings($"marginTop={PointsInCm(pageMargin.Top)}cm");
      WReport.WriteSettings($"marginBottom={PointsInCm(pageMargin.Bottom)}cm");
      WReport.WriteSettings($"marginLeft={PointsInCm((int)pageMargin.Left.Value)}cm");
      WReport.WriteSettings($"marginRight={PointsInCm((int)pageMargin.Right.Value)}cm");
      WReport.WriteSettings($"marginHeader={PointsInCm((int)pageMargin.Header.Value)}cm");
      WReport.WriteSettings($"marginFooter={PointsInCm((int)pageMargin.Footer.Value)}cm");
      WReport.WriteSettings("");
    }
  }
}