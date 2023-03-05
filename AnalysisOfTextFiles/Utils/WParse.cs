using System.Collections.Generic;
using System.Linq;
using AnalysisOfTextFiles.Objects;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
public class WParse
{
  public static void Content()
  {
    State.Styles = WStyles.GetDocStyles();
    WReport.CreateReportFile();

    _Header();
    _Footer();
    _Body();
    
    State.WDocument.Close();
  }
  private static void _Body()
  {
    Body body = State.WDocument.MainDocumentPart.Document.Body;
    
    List<Paragraph> descendants = body.Descendants<Paragraph>().ToList();
    foreach (Paragraph parDesc in descendants)
    {
      int idx = descendants.IndexOf(parDesc);
      if (((string)parDesc.ParagraphProperties?.ParagraphStyleId?.Val)?.StartsWith("TOC") == true)
      {
        // This is a TOC entry
        Analis.ParagraphCheck(parDesc, idx, Analis.ContentType.TOC);
      }
      else
      {
        // Check if the paragraph is part of a table
        if (parDesc.Parent != null && parDesc.Parent is TableCell)
        {
          TableCell cell = (TableCell)parDesc.Parent;
          TableRow row = (TableRow)parDesc.Parent.Parent;
          Table table = (Table)parDesc.Parent.Parent.Parent;

          int parIdx = cell.Descendants<Paragraph>().ToList().IndexOf(parDesc);
          int cellIdx = row.Descendants<TableCell>().ToList().IndexOf(cell);
          int rowIdx = table.Descendants<TableRow>().ToList().IndexOf(row);
          int tableIdx = body.Descendants<Table>().ToList().IndexOf(table);
          
          WTable Wtable = new WTable(tableIdx, rowIdx, cellIdx, parIdx);
          
          // This is a table row
          Analis.ParagraphCheck(parDesc, idx, Analis.ContentType.Table, Wtable);
        }
        else
        {
          // This is an ordinary paragraph
          Analis.ParagraphCheck(parDesc, idx, Analis.ContentType.Paragraph);
        }
      }
    }
  }
  private static void _Header()
  {
    List<HeaderPart> headerParts = State.WDocument.MainDocumentPart.HeaderParts.ToList();

    foreach (HeaderPart headerPart in headerParts)
    {
      Header headers = headerPart.Header;
      List<Paragraph> paragraphs = headers.Descendants<Paragraph>().ToList();
      foreach (var paragraph in paragraphs)
      {
        int idx = paragraphs.IndexOf(paragraph);
        Analis.ParagraphCheck(paragraph, idx, Analis.ContentType.Header);
      }
    }
  }
  private static void _Footer()
  {
    List<FooterPart> footerParts = State.WDocument.MainDocumentPart.FooterParts.ToList();
    
    foreach (FooterPart footerPart in footerParts)
    {
      Footer footers = footerPart.Footer;
      List<Paragraph> paragraphs = footers.Descendants<Paragraph>().ToList();
      foreach (var paragraph in paragraphs)
      {
        int idx = paragraphs.IndexOf(paragraph);
        Analis.ParagraphCheck(paragraph, idx, Analis.ContentType.Footer);
      }
    }
  }
}