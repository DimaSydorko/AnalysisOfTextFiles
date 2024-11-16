using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Linq;

namespace AnalysisOfTextFiles.Objects;

public class Order
{
  public enum OrderCheckType
  {
    Before,
    After
  }
    public static void CheckParagraph(Paragraph paragraph, Paragraph prevParagraph, Paragraph nextParagraph, CheckParagraph.ContentType type, string styleName, int idx)
  {
    StyleProperties? styleSettings = State.StylesSettings.FirstOrDefault(s => s.name == styleName);

    if (styleSettings?.before != null && styleSettings.before.Count > 0)
    {
      if (nextParagraph != null)
      {
        _CheckParagraphDependence(paragraph, nextParagraph, idx, type, styleSettings.before,
          OrderCheckType.Before);
      }
      else
      {
        WReport.Order(paragraph, type, idx, styleName, styleSettings.before, WReport.OrderType.MissedBefore);
      }
    }

    else if (styleSettings?.after != null && styleSettings.after.Count > 0)
    {
      if (prevParagraph != null)
      {
        _CheckParagraphDependence(paragraph, prevParagraph, idx, type, styleSettings.after, OrderCheckType.After);
      }
      else
      {
        WReport.Order(paragraph, type, idx, styleName, styleSettings.after, WReport.OrderType.MissedAfter);
      }
    }
  }

  private static void _CheckParagraphDependence(Paragraph paragraph, Paragraph reviewParagraph, int idx,
    CheckParagraph.ContentType type, List<string> depends, OrderCheckType orderCheckType)
  {
    string? styleName = WDecoding.RemoveSuffixIfExists(Objects.CheckParagraph.GetParagraphStyle(paragraph));
    string? reviewStyleName = WDecoding.RemoveSuffixIfExists(Objects.CheckParagraph.GetParagraphStyle(reviewParagraph));

    WReport.OrderType orderType = orderCheckType switch
    {
      OrderCheckType.After => WReport.OrderType.InsteadAfter,
      _ => WReport.OrderType.InsteadBefore
    };

    bool isParaEmpty = reviewParagraph.ParagraphProperties == null;
    if (isParaEmpty)
    {
      orderType = orderCheckType switch
      {
        OrderCheckType.After => WReport.OrderType.MissedAfter,
        _ => WReport.OrderType.MissedBefore
      };

      WReport.Order(paragraph, type, idx, styleName, depends, orderType);
      return;
    }

    void onOrderReport(WTable? table = null)
    {
      if (!depends.Contains(reviewStyleName ?? ""))
      {
        WReport.Order(reviewParagraph, type, idx, styleName, depends, orderType, reviewStyleName, table);
      }
    }

    if (reviewParagraph.Parent?.LocalName == "sdtContent") onOrderReport();
    else if (reviewParagraph.Parent is TableCell)
    {
      TableCell cell = (TableCell)reviewParagraph.Parent;
      TableRow row = (TableRow)reviewParagraph.Parent.Parent;
      Table table = (Table)reviewParagraph.Parent.Parent.Parent;

      int parIdx = cell.Descendants<Paragraph>().ToList().IndexOf(reviewParagraph);
      int cellIdx = row.Descendants<TableCell>().ToList().IndexOf(cell);
      int rowIdx = table.Descendants<TableRow>().ToList().IndexOf(row);
      int tableIdx = reviewParagraph.Parent.Parent.Parent.Parent.Descendants<Table>().ToList().IndexOf(table);
      WTable Wtable = new WTable(tableIdx, rowIdx, cellIdx, parIdx);

      onOrderReport(Wtable);
    }
    else onOrderReport();
  }
}