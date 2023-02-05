using System.Collections.Generic;
using System.Linq;
using System.Windows;
using AnalysisOfTextFiles.Objects;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;

namespace AnalysisOfTextFiles
{
  /// <summary>
  /// Interaction logic for MainWindow.xaml
  /// </summary>
  public partial class MainWindow : Window
  {
    public MainWindow()
    {
      InitializeComponent();
    }

    private void Upload_OnClick(object sender, RoutedEventArgs e)
    {
      string[] allowedStyles = { "Heading1", "Heading2", "Heading3" };

      WFileName fileName = new WFileName();
      fileName.Open();
      //Open and clone file                                                                       
      using WordprocessingDocument sourceWordDocument = WordprocessingDocument.Open(fileName.full, false);
      using WordprocessingDocument document = (WordprocessingDocument)sourceWordDocument.Clone(fileName.analized, true);

      // Assign a reference to the appropriate part to the stylesPart variable.
      Styles stylesXml = document.MainDocumentPart.StyleDefinitionsPart.Styles;
      List<WStyle> docStyles = Analis.ExtractStyles(stylesXml);
      
      Body body = document.MainDocumentPart.Document.Body;
      foreach (Paragraph para in body.Elements<Paragraph>())
      {
        ParagraphProperties pPr = para.GetFirstChild<ParagraphProperties>();

        if (pPr == null || pPr.GetFirstChild<ParagraphStyleId>() == null) continue;
        
        if (pPr == null || pPr.GetFirstChild<ParagraphStyleId>() == null)
        {    
          if (!string.IsNullOrEmpty(para.InnerText))
          {       
            WComment.Add(document.MainDocumentPart, para, "Normal");
          }                 
        }
        else
        {
          string pStyleEncoded = pPr.GetFirstChild<ParagraphStyleId>().Val;

          WStyle style = docStyles.SingleOrDefault(s => { return s.encoded == pStyleEncoded; });
          string first4Letters = style.decoded.Substring(0, 4);

          // if the value of the pStyle is allowed => skip the paragraph
          if (allowedStyles.Contains(pStyleEncoded) || first4Letters == "ЕОМ:") continue;

          WComment.Add(document.MainDocumentPart, para, style.decoded);
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