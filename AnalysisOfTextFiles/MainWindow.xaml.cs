using System;
using System.Collections.Generic;
using System.Windows;
using AnalysisOfTextFiles.Objects;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System.Linq;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;

namespace AnalysisOfTextFiles
{
  public partial class MainWindow
  {
    public MainWindow()
    {
      InitializeComponent();
    }

    private static bool IsСomments { get; set; } = true;

    private void Upload_OnClick(object sender, RoutedEventArgs e)
    {
      WFileName fileName = new WFileName();
      fileName.Open();
      //Open and clone file                                                                       
      using WordprocessingDocument sourceWordDocument = WordprocessingDocument.Open(fileName.full, false);
      using WordprocessingDocument document = IsСomments
        ? (WordprocessingDocument)sourceWordDocument.Clone(fileName.analized, true)
        : sourceWordDocument;
      string timestamp = DateTime.Now.ToString("F");
      File.WriteAllText(fileName.report, $"-----------------Report ({timestamp})----------------\n");

      // Assign a reference to the appropriate part to the stylesPart variable.
      MainDocumentPart mainPart = document.MainDocumentPart;
      Styles stylesXml = mainPart.StyleDefinitionsPart.Styles;
      Body body = mainPart.Document.Body;

      List<WStyle> docStyles = Analis.ExtractStyles(stylesXml);

      Analis.Set(IsСomments, docStyles, fileName.report);

      //--------header-------- 
      foreach (HeaderPart headerPart in mainPart.HeaderParts.ToList())
      {
        Header headers = headerPart.Header;
        List<Paragraph> paragraphs = headers.Descendants<Paragraph>().ToList();
        foreach (var paragraph in paragraphs)
        {
          int idx = paragraphs.IndexOf(paragraph);
          Analis.ParagraphCheck(paragraph, mainPart, idx, Analis.ContentType.Header);
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
          Analis.ParagraphCheck(paragraph, mainPart, idx, Analis.ContentType.Footer);
        }
      }
      
      // //--------paragraph-------- 
      // List<Paragraph> bodyParagraphs = body.Elements<Paragraph>().ToList();
      // foreach (Paragraph paragraph in bodyParagraphs)
      // {
      //   int idx = bodyParagraphs.IndexOf(paragraph);
      //   Analis.ParagraphCheck(paragraph, mainPart, idx, Analis.ContentType.Paragraph);
      // }

      // //--------tables-------- 
      // List<Table> tables = body.Descendants<Table>().ToList();
      // foreach (Table table in tables)
      // {
      //   int idx = tables.IndexOf(table);
      //   IEnumerable<TableCell> cells = table.Descendants<TableCell>();
      //   foreach (TableCell cell in cells)
      //   {
      //     IEnumerable<Paragraph> paragraphs = cell.Descendants<Paragraph>();
      //     foreach (Paragraph paragraph in paragraphs)
      //     {
      //       Analis.ParagraphCheck(paragraph, mainPart, idx, Analis.ContentType.Table);
      //     }
      //   }
      // }
      
      //--------File parsing-------- 
      List<Paragraph> descendants = body.Descendants<Paragraph>().ToList();
      foreach (Paragraph parDesc in descendants)
      {
        int idx = descendants.IndexOf(parDesc);
        if (parDesc.ParagraphProperties != null && parDesc.ParagraphProperties.ParagraphStyleId != null
                                                && parDesc.ParagraphProperties.ParagraphStyleId.Val.HasValue
                                                && parDesc.ParagraphProperties.ParagraphStyleId.Val.Value.StartsWith(
                                                  "TOC"))
        {
          // This is a TOC entry
          Analis.ParagraphCheck(parDesc, mainPart, idx, Analis.ContentType.TOC);
        }
        else
        {
          // Check if the paragraph is part of a table
          if (parDesc.Parent != null && parDesc.Parent.LocalName == "tc")
          {
            // This is a table row
            Analis.ParagraphCheck(parDesc, mainPart, idx, Analis.ContentType.Table);
          }
          else
          {
            // This is an ordinary paragraph
            Analis.ParagraphCheck(parDesc, mainPart, idx, Analis.ContentType.Paragraph);
          }
        }
      }
      
      MessageBox.Show($"File {fileName.withoutExtension} analysed", "Complete Status");
    }

    private void RewriteCheckBox_OnClick(object sender, RoutedEventArgs e)
    {
      IsСomments = !IsСomments;
    }
  }
}