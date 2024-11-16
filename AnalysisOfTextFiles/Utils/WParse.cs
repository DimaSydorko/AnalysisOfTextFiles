using System.Linq;
using AnalysisOfTextFiles.Objects;
using DocumentFormat.OpenXml.Wordprocessing;

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
    CheckPage.AnalisePageSettings();
    WStyles.AnaliseStylesSettings();
    State.WDocument.Close();
  }

  private static void _Body()
  {
    WReport.Write("----Boody----");

    var body = State.WDocument.MainDocumentPart.Document.Body;

    var descendants = body.Descendants<Paragraph>().ToList();

    void AnalyzeParagraph(Paragraph parDesc)
    {
      var idx = descendants.IndexOf(parDesc);
      Paragraph nextParagraph = idx == (descendants.Count - 1) ? null : descendants[idx + 1];
      Paragraph prevParagraph = idx == 0 ? null : descendants[idx - 1];

      string prev = WDecoding.RemoveSuffixIfExists(CheckParagraph.GetParagraphStyle(prevParagraph));
      string curr = WDecoding.RemoveSuffixIfExists(CheckParagraph.GetParagraphStyle(parDesc));
      string next = WDecoding.RemoveSuffixIfExists(CheckParagraph.GetParagraphStyle(nextParagraph));

      if (parDesc.Parent.LocalName == "sdtContent")
        CheckParagraph.ParagraphCheck(parDesc, prevParagraph, nextParagraph, idx, CheckParagraph.ContentType.TOC);
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

        CheckParagraph.ParagraphCheck(parDesc, prevParagraph, nextParagraph, idx, CheckParagraph.ContentType.Table,
          Wtable);
      }
      else
        CheckParagraph.ParagraphCheck(parDesc, prevParagraph, nextParagraph, idx, CheckParagraph.ContentType.Paragraph);
    }

    foreach (var parDesc in descendants)
    {
      AnalyzeParagraph(parDesc);
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
        CheckParagraph.ParagraphCheck(paragraph, null, null, idx, CheckParagraph.ContentType.Header);
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
        CheckParagraph.ParagraphCheck(paragraph, null, null, idx, CheckParagraph.ContentType.Footer);
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