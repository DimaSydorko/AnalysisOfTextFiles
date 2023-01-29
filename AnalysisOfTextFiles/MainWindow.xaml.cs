using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Win32;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;

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

      string fileName;
      OpenFileDialog openFileDialog = new OpenFileDialog();
      openFileDialog.Filter = "*.doc|*.docx";
      openFileDialog.InitialDirectory = @"c:\temp\";
      if (openFileDialog.ShowDialog() == true) fileName = openFileDialog.FileName;
      else return;

      string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(fileName);
      string directoryName = Path.GetDirectoryName(fileName);

      //Open and clone file
      using WordprocessingDocument sourceWordDocument = WordprocessingDocument.Open(fileName, false);
      using WordprocessingDocument document =
        (WordprocessingDocument)sourceWordDocument.Clone($"{directoryName}/{fileNameWithoutExtension} ANALYSED.docx",
          true);

      Body body = document.MainDocumentPart.Document.Body;

      // Assign a reference to the appropriate part to the stylesPart variable.
      Styles stylesXml = document.MainDocumentPart.StyleDefinitionsPart.Styles;
      WStyle[] docStyles = Analis.ExtractStyles(stylesXml);

      // MyDocuments.Body is a WordProcessDocument.MainDocumentPart.Document.Body
      foreach (Paragraph para in body.Elements<Paragraph>())
      {
        // if the paragraph has no properties or has properties but no pStyle, it's not a "Heading1"
        ParagraphProperties pPr = para.GetFirstChild<ParagraphProperties>();
        if (pPr == null || pPr.GetFirstChild<ParagraphStyleId>() == null) continue;

        // if the value of the pStyle is allowed => skip the paragraph
        string pStyle = pPr.GetFirstChild<ParagraphStyleId>().Val;

        // var style = docStyles.SingleOrDefault(s => { return s.decoded == pStyle; });
        if (allowedStyles.Contains(pStyle) || pStyle.Substring(0, 3) == "EOM") continue;

        // MessageBox.Show($"{pStyle.Substring(0, 3)}_{pStyle.Substring(0, 3).Equals("ЕОМ")}");

        int id = 0;
        Comments comments;

        // Verify that the document contains a
        // WordProcessingCommentsPart part; if not, add a new one.
        if (document.MainDocumentPart.GetPartsCountOfType<WordprocessingCommentsPart>() > 0)
        {
          comments =
            document.MainDocumentPart.WordprocessingCommentsPart.Comments;
          if (comments.HasChildren)
          {
            // Obtain an unused ID.
            id = Int32.Parse(comments.Descendants<Comment>().Select(e => e.Id.Value).Max()) + 1;
          }
        }
        else
        {
          // No WordprocessingCommentsPart part exists, so add one to the package.
          WordprocessingCommentsPart commentPart =
            document.MainDocumentPart.AddNewPart<WordprocessingCommentsPart>();
          commentPart.Comments = new Comments();
          comments = commentPart.Comments;
        }

        // Compose a new Comment and add it to the Comments part.
        Paragraph par = new Paragraph(new Run(new Text(pPr.GetFirstChild<ParagraphStyleId>().Val)));
        Comment cmt =
          new Comment()
          {
            Id = id.ToString(),
            Author = "Wrong style",
          };
        cmt.AppendChild(par);
        comments.AppendChild(cmt);
        comments.Save();

        // Specify the text range for the Comment.
        // Insert the new CommentRangeStart before the first run of paragraph.
        para.InsertBefore(new CommentRangeStart() { Id = id.ToString() }, para.GetFirstChild<Run>());

        // Insert the new CommentRangeEnd after last run of paragraph.
        var cmtEnd = para.InsertAfter(new CommentRangeEnd() { Id = id.ToString() }, para.Elements<Run>().Last());

        // Compose a run with CommentReference and insert it.
        para.InsertAfter(new Run(new CommentReference() { Id = id.ToString() }), cmtEnd);
      }

      MessageBox.Show($"File {fileNameWithoutExtension} analysed", "Complete Status");
    }

    private void StyleSettings_OnClick(object sender, RoutedEventArgs e)
    {
      MessageBox.Show("StyleSettings_OnClick");
    }
  }
}

class WStyle
{
  public string decoded { get; set; }
  public string encoded { get; set; }

  public void SetDec(string dec)
  {
    decoded = dec;
  }

  public void SetEnc(string enc)
  {
    encoded = enc;
  }
}

class Analis
{
  public static WStyle[] ExtractStyles(Styles stylesXml)
  {
    WStyle[] styles = new WStyle[2];
    // WStyle[] styles = Array.Empty<WStyle>();

    //Map to get Encoded and Decoded StyleNames
    foreach (var styleXml in stylesXml.ChildElements)
    {
      if (styleXml.ChildElements.Count >= 3)
      {
        var styleDec = styleXml.ChildElements[0];
        var styleEnc = styleXml.ChildElements[2];

        WStyle style = new WStyle();

        //Get Decoded Name
        PropertyInfo propertyDec = styleDec.GetType().GetProperty("Val");
        if (propertyDec != null)
        {
          var styleNameObjDec = propertyDec.GetValue(styleDec);
          PropertyInfo propertyDecName = styleNameObjDec.GetType().GetProperty("Value");
          string decName = propertyDecName.GetValue(styleNameObjDec).ToString();

          style.SetDec(decName);
        }

        //Get Encoded Name
        PropertyInfo propertyEnc = styleEnc.GetType().GetProperty("Val");
        if (propertyEnc != null)
        {
          var styleNameObjEnc = propertyEnc.GetValue(styleEnc);
          if (styleNameObjEnc != null)
          {
            PropertyInfo propertyEncName = styleNameObjEnc.GetType().GetProperty("Value");
            string decName = propertyEncName.GetValue(styleNameObjEnc).ToString();

            style.SetEnc(decName);
          }
        }

        if (style.decoded != null)
        {
          styles.Append(style);
        }
      }
    }
          //TODO Fix Save elements to array
    return styles;
  }
}