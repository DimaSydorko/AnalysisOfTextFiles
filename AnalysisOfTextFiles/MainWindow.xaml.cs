using System;
using System.Linq;
using System.Windows;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Win32;
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

    private bool IsRewriteComments { get; set; } = true;

    private void Upload_OnClick(object sender, RoutedEventArgs e)
    {
      string fileName;
      OpenFileDialog openFileDialog = new OpenFileDialog();
      openFileDialog.Filter = "*.doc|*.docx";
      openFileDialog.InitialDirectory = @"c:\temp\";
      if (openFileDialog.ShowDialog() == true) fileName = openFileDialog.FileName;
      else return;

      using (WordprocessingDocument myDocument = WordprocessingDocument.Open(fileName, true))
      {
        Body body = myDocument.MainDocumentPart.Document.Body;

        // MyDocuments.Body is a WordProcessDocument.MainDocumentPart.Document.Body
        foreach (Paragraph para in body.Elements<Paragraph>())
        {
          // if the paragraph has no properties or has properties but no pStyle, it's not a "Heading1"
          ParagraphProperties pPr = para.GetFirstChild<ParagraphProperties>();
          if (pPr == null || pPr.GetFirstChild<ParagraphStyleId>() == null) continue;
          // if the value of the pStyle is Heading3 => skip the paragraph
          if (pPr.GetFirstChild<ParagraphStyleId>().Val == "Heading3") continue;

          int id = 0;
          Comments comments;

          // Verify that the document contains a
          // WordProcessingCommentsPart part; if not, add a new one.
          if (myDocument.MainDocumentPart.GetPartsCountOfType<WordprocessingCommentsPart>() > 0)
          {
            // if (IsRewriteComments)
            // {
              // myDocument.MainDocumentPart.WordprocessingCommentsPart.Comments = new Comments();
            // }
            // else
            // {
              comments = myDocument.MainDocumentPart.WordprocessingCommentsPart.Comments;

              if (comments.HasChildren)
              {
                // Obtain an unused ID.
                id = Int32.Parse(comments.Descendants<Comment>().Select(e => e.Id.Value).Max()) + 1;
              }
            }
          // }
          else
          {
            // if (IsRewriteComments) myDocument.MainDocumentPart.DeletePart(myDocument.MainDocumentPart.WordprocessingCommentsPart);

            // No WordprocessingCommentsPart part exists, so add one to the package.
            WordprocessingCommentsPart commentPart =
              myDocument.MainDocumentPart.AddNewPart<WordprocessingCommentsPart>();
            commentPart.Comments = new Comments();
            comments = commentPart.Comments;
          }

          // Compose a new Comment and add it to the Comments part.
          Paragraph par =
            new Paragraph(
              new Run(new Text(pPr.GetFirstChild<ParagraphStyleId>().Val)));
          Comment cmt = new Comment()
          {
            Id = id.ToString(),
            Author = "Wrong style",
            Date = DateTime.Now.ToLocalTime(),
          };
          cmt.AppendChild(par);
          comments.AppendChild(cmt);
          comments.Save();

          // Specify the text range for the Comment.
          // Insert the new CommentRangeStart before the first run of paragraph.
          para.InsertBefore(new CommentRangeStart() { Id = id.ToString() },
            para.GetFirstChild<Run>());

          // Insert the new CommentRangeEnd after last run of paragraph.
          var cmtEnd = para.InsertAfter(new CommentRangeEnd() { Id = id.ToString() },
            para.Elements<Run>().Last());

          // Compose a run with CommentReference and insert it.
          para.InsertAfter(new Run(new CommentReference() { Id = id.ToString() }), cmtEnd);
        }

        MessageBox.Show($"File {fileName} analysed", "Complete Status");
      }
    }

    private void StyleSettings_OnClick(object sender, RoutedEventArgs e)
    {
      MessageBox.Show($"isRewriteComments {IsRewriteComments}");
      MessageBox.Show("StyleSettings_OnClick");
    }

    private void RewriteCheckBox_OnClick(object sender, RoutedEventArgs e)
    {
      IsRewriteComments = !IsRewriteComments;
    }
  }
}