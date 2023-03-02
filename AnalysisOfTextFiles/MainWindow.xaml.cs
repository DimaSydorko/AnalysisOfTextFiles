using System;
using System.Collections.Generic;
using System.Windows;
using AnalysisOfTextFiles.Objects;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;

using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;

namespace AnalysisOfTextFiles
{
  public partial class MainWindow
  {
    public MainWindow()
    {
      InitializeComponent();
    }

    private void Upload_OnClick(object sender, RoutedEventArgs e)
    {
      WFileName fileName = new WFileName();
      fileName.Open();
      //Open and clone file                                                                       
      using WordprocessingDocument sourceWordDocument = WordprocessingDocument.Open(fileName.full, false);
      using WordprocessingDocument document = (WordprocessingDocument)sourceWordDocument.Clone(fileName.analized, true);
      string timestamp = DateTime.Now.ToString("F");
      File.WriteAllText(fileName.report, $"-----------------Report ({timestamp})----------------\n");

      // Assign a reference to the appropriate part to the stylesPart variable.
      MainDocumentPart mainPart = document.MainDocumentPart;
      Styles stylesXml = mainPart.StyleDefinitionsPart.Styles;
      Body body = mainPart.Document.Body;

      List<WStyle> docStyles = Analis.ExtractStyles(stylesXml);

      //--------header-------- 
      foreach (HeaderPart headerPart in mainPart.HeaderParts)
      {
        Header headers = headerPart.Header;
        List<Paragraph> paragraphs = headers.Descendants<Paragraph>().ToList();
        foreach (var paragraph in paragraphs)
        {
          int idx = paragraphs.IndexOf(paragraph);
          Analis.ParagraphCheck(paragraph, mainPart, docStyles, fileName.report, idx, Analis.ContentType.Header);
        }
      }

      //--------footer-------- 
      foreach (FooterPart footerPart in mainPart.FooterParts)
      {
        Footer footers = footerPart.Footer;
        List<Paragraph> paragraphs = footers.Descendants<Paragraph>().ToList();
        foreach (var paragraph in paragraphs)
        {
          int idx = paragraphs.IndexOf(paragraph);
          Analis.ParagraphCheck(paragraph, mainPart, docStyles, fileName.report, idx, Analis.ContentType.Footer);
        }
      }

      //--------paragraph-------- 
      List<Paragraph> bodyParagraphs = body.Elements<Paragraph>().ToList();
      foreach (Paragraph paragraph in bodyParagraphs)
      {
        int idx = bodyParagraphs.IndexOf(paragraph);
        Analis.ParagraphCheck(paragraph, mainPart, docStyles, fileName.report, idx, Analis.ContentType.Paragraph);
      }

      //--------tables-------- 
      List<Table> tables = body.Descendants<Table>().ToList();
      foreach (Table table in tables)
      {
        int idx = tables.IndexOf(table);
        IEnumerable<TableCell> cells = table.Descendants<TableCell>();
        foreach (TableCell cell in cells)
        {
          IEnumerable<Paragraph> paragraphs = cell.Descendants<Paragraph>();
          foreach (Paragraph paragraph in paragraphs)
          {
            Analis.ParagraphCheck(paragraph, mainPart, docStyles, fileName.report, idx, Analis.ContentType.Table);
          }
        }
      }

      //--------table of content-------- 
      OpenXmlElement toc = body.Descendants().FirstOrDefault();
      if (toc != null)
      {
        List<Paragraph> tocEls = toc.Descendants<Paragraph>().ToList();
        foreach (Paragraph paragraph in tocEls)
        {
          Analis.ParagraphCheck(paragraph, mainPart, docStyles, fileName.report, 0, Analis.ContentType.TOC);
        }
      }

      MessageBox.Show($"File {fileName.withoutExtension} analysed", "Complete Status");
    }

    private void StyleSettings_OnClick(object sender, RoutedEventArgs e)
    {
      MessageBox.Show("StyleSettings_OnClick");
    }
  }
}