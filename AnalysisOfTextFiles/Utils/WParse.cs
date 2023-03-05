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
    // Assign a reference to the appropriate part to the stylesPart variable.
    MainDocumentPart mainPart = State.WDocument.MainDocumentPart;
    Body body = mainPart.Document.Body;

    State.Styles = WStyles.GetDocStyles();
    WReport.CreateReportFile();

    //--------header-------- 
    foreach (HeaderPart headerPart in mainPart.HeaderParts.ToList())
    {
      Header headers = headerPart.Header;
      List<Paragraph> paragraphs = headers.Descendants<Paragraph>().ToList();
      foreach (var paragraph in paragraphs)
      {
        int idx = paragraphs.IndexOf(paragraph);
        Analis.ParagraphCheck(paragraph, idx, Analis.ContentType.Header);
      }
    }

    //--------footer-------- 
    foreach (FooterPart footerPart in mainPart.FooterParts.ToList())
    {
      Footer footers = footerPart.Footer;
      List<Paragraph> paragraphs = footers.Descendants<Paragraph>().ToList();
      foreach (var paragraph in paragraphs)
      {
        int idx = paragraphs.IndexOf(paragraph);
        Analis.ParagraphCheck(paragraph, idx, Analis.ContentType.Footer);
      }
    }

    //--------File parsing-------- 
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
    
    State.WDocument.Close();
  }
}