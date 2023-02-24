using System.Collections.Generic;
using System.Windows;
using AnalysisOfTextFiles.Objects;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
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

      // Assign a reference to the appropriate part to the stylesPart variable.
      MainDocumentPart mainPart = document.MainDocumentPart;
      Styles stylesXml = mainPart.StyleDefinitionsPart.Styles;

      List<WStyle> docStyles = Analis.ExtractStyles(stylesXml);


      Body body = mainPart.Document.Body;
      //--------paragraph-------- 
      foreach (Paragraph paragraph in body.Elements<Paragraph>())
      {
        Analis.ParagraphCheck(paragraph, mainPart, docStyles);
      }

      //--------tables-------- 
      IEnumerable<Table> tables = body.Descendants<Table>();
      foreach (Table table in tables)
      {
        IEnumerable<TableCell> cells = table.Descendants<TableCell>();
        foreach (TableCell cell in cells)
        {
          IEnumerable<Paragraph> paragraphs = cell.Descendants<Paragraph>();
          foreach (Paragraph paragraph in paragraphs)
          {
            Analis.ParagraphCheck(paragraph,mainPart, docStyles);
          }
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