using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AnalysisOfTextFiles.Objects;

public class Order
{
  public enum OrderCheckType
  {
    Before,
    After
  }

  public static void CheckParagraph(Paragraph paragraph, CheckParagraph.ContentType type, string styleName, int idx,
    WTable? table)
  {
    var styleSettings = State.StylesSettings.FirstOrDefault(s => s.name == styleName);

    if (styleSettings?.before != null && styleSettings.before.Count > 0)
    {
      if (State.NextParagraphName != null)
        _CheckParagraphDependence(paragraph, State.NextParagraphName, idx, type, styleSettings.before,
          OrderCheckType.Before, table);
      else
        WReport.Order(paragraph, type, idx, styleName, styleSettings.before, WReport.OrderType.MissedBefore);
    }

    else if (styleSettings?.after != null && styleSettings.after.Count > 0)
    {
      if (State.PrevParagraphName != null)
        _CheckParagraphDependence(paragraph, State.PrevParagraphName, idx, type, styleSettings.after,
          OrderCheckType.After, table);
      else
        WReport.Order(paragraph, type, idx, styleName, styleSettings.after, WReport.OrderType.MissedAfter);
    }
  }

  private static void _CheckParagraphDependence(Paragraph paragraph, string reviewStyleName, int idx,
    CheckParagraph.ContentType type, List<string> depends, OrderCheckType orderCheckType, WTable? table)
  {
    var styleName = WDecoding.RemoveSuffixIfExists(Objects.CheckParagraph.GetParagraphStyle(paragraph));

    var orderType = orderCheckType switch
    {
      OrderCheckType.After => WReport.OrderType.InsteadAfter,
      _ => WReport.OrderType.InsteadBefore
    };
    if (!depends.Contains(reviewStyleName ?? ""))
      WReport.Order(paragraph, type, idx, styleName, depends, orderType, reviewStyleName, table);
  }
}