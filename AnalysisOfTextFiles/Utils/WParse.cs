using System.Linq;
using AnalysisOfTextFiles.Objects;
using DocumentFormat.OpenXml.Wordprocessing;
using Toc = DocumentFormat.OpenXml.Wordprocessing.Table;

public class WParse
{
  public static void Content()
  {
    WReport.CreateReportFile();
    State.Styles = WStyles.GetDocStyles();
    WStyles.Review();

    _Header();
    _Footer();
    _Body();

    State.WDocument.Close();
  }

  private static void _Body()
  {
    var body = State.WDocument.MainDocumentPart.Document.Body;

    var descendants = body.Descendants<Paragraph>().ToList();
    foreach (var parDesc in descendants)
    {
      var idx = descendants.IndexOf(parDesc);
      if (parDesc.Parent.LocalName == "sdtContent")
      {
        // This is a TOC entry
        Analis.ParagraphCheck(parDesc, idx, Analis.ContentType.TOC);
      }
      else
      {
        // Check if the paragraph is part of a table
        if (parDesc.Parent != null && parDesc.Parent is TableCell)
        {
          var cell = (TableCell)parDesc.Parent;
          var row = (TableRow)parDesc.Parent.Parent;
          var table = (Table)parDesc.Parent.Parent.Parent;

          var parIdx = cell.Descendants<Paragraph>().ToList().IndexOf(parDesc);
          var cellIdx = row.Descendants<TableCell>().ToList().IndexOf(cell);
          var rowIdx = table.Descendants<TableRow>().ToList().IndexOf(row);
          var tableIdx = body.Descendants<Table>().ToList().IndexOf(table);

          var Wtable = new WTable(tableIdx, rowIdx, cellIdx, parIdx);

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
    var headerParts = State.WDocument.MainDocumentPart.HeaderParts.ToList();

    foreach (var headerPart in headerParts)
    {
      var headers = headerPart.Header;
      var paragraphs = headers.Descendants<Paragraph>().ToList();
      foreach (var paragraph in paragraphs)
      {
        var idx = paragraphs.IndexOf(paragraph);
        Analis.ParagraphCheck(paragraph, idx, Analis.ContentType.Header);
      }
    }
  }

  private static void _Footer()
  {
    var footerParts = State.WDocument.MainDocumentPart.FooterParts.ToList();

    foreach (var footerPart in footerParts)
    {
      var footers = footerPart.Footer;
      var paragraphs = footers.Descendants<Paragraph>().ToList();
      foreach (var paragraph in paragraphs)
      {
        var idx = paragraphs.IndexOf(paragraph);
        Analis.ParagraphCheck(paragraph, idx, Analis.ContentType.Footer);
      }
    }
  }
}