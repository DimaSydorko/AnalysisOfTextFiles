using System.Collections.Generic;

namespace AnalysisOfTextFiles.Objects;

public class WPage
{
  public WPage(List<string> size, List<string> orientation, double marginTop, double marginBottom, double marginLeft,
    double marginRight, double marginHeader, double marginFooter)
  {
    Size = size;
    Orientation = orientation;
    MarginTop = marginTop;
    MarginBottom = marginBottom;
    MarginLeft = marginLeft;
    MarginRight = marginRight;
    MarginHeader = marginHeader;
    MarginFooter = marginFooter;
  }

  public List<string> Size { get; }
  public List<string> Orientation { get; }
  public double MarginTop { get; }
  public double MarginBottom { get; }
  public double MarginLeft { get; }
  public double MarginRight { get; }
  public double MarginHeader { get; }
  public double MarginFooter { get; }
}