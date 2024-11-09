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

    _PageFormat();
    _Header();
    _Footer();
    _Body();

    State.WDocument.Close();
  }

  public static void StyleSettings()
  {
    WReport.CreateStylesFile();
    State.Styles = WStyles.GetDocStyles();
    WStyles.AnaliseStylesSettings();
    State.WDocument.Close();
  }

  private static void _Body()
  {
    WReport.Write("----Boody----");

    var body = State.WDocument.MainDocumentPart.Document.Body;

    var descendants = body.Descendants<Paragraph>().ToList();
    foreach (var parDesc in descendants)
    {
      var idx = descendants.IndexOf(parDesc);
      State.NextParagraph = idx == (descendants.Count - 1) ? null : descendants[idx + 1];
      State.PrevParagraph = idx == 0 ? null : descendants[idx - 1];

      string prev = WDecoding.RemoveSuffixIfExists(CheckParagraph.GetParagraphStyle(State.PrevParagraph));
      string curr = WDecoding.RemoveSuffixIfExists(CheckParagraph.GetParagraphStyle(parDesc));
      string next = WDecoding.RemoveSuffixIfExists(CheckParagraph.GetParagraphStyle(State.NextParagraph));

      // Check if it's TOC
      if (parDesc.Parent.LocalName == "sdtContent") CheckParagraph.ParagraphCheck(parDesc, idx, CheckParagraph.ContentType.TOC);
      else if (parDesc.Parent != null && parDesc.Parent is TableCell)
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
        CheckParagraph.ParagraphCheck(parDesc, idx, CheckParagraph.ContentType.Table, Wtable);
        // This is an ordinary paragraph
      }
      else CheckParagraph.ParagraphCheck(parDesc, idx, CheckParagraph.ContentType.Paragraph);
    }
  }

  private static void _Header()
  {
    WReport.Write("----Header & Footer----");

    var headerParts = State.WDocument.MainDocumentPart.HeaderParts.ToList();

    foreach (var headerPart in headerParts)
    {
      var headers = headerPart.Header;
      var paragraphs = headers.Descendants<Paragraph>().ToList();
      foreach (var paragraph in paragraphs)
      {
        var idx = paragraphs.IndexOf(paragraph);
        CheckParagraph.ParagraphCheck(paragraph, idx, CheckParagraph.ContentType.Header);
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
        CheckParagraph.ParagraphCheck(paragraph, idx, CheckParagraph.ContentType.Footer);
      }
    }
  }

  private static void _PageFormat()
  {
    WReport.Write("----Page format----");

    var body = State.WDocument.MainDocumentPart.Document.Body;

    foreach (var section in body.Elements<SectionProperties>())
    {
      CheckPage.CheckDimensions(section);
      CheckPage.CheckPageMargin(section);
    }
  }
}