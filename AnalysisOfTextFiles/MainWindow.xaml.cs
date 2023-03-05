using System;
using System.Windows;
using AnalysisOfTextFiles.Objects;
using DocumentFormat.OpenXml.Packaging;

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
      State.FilePath = WFilePath.Open();
      State.IsСomments = IsСomments;

      WordprocessingDocument document = null;
      try
      {
        //Open and clone file                                                                       
        using WordprocessingDocument sourceWordDocument = WordprocessingDocument.Open(State.FilePath.full, false);
        document = IsСomments
          ? (WordprocessingDocument)sourceWordDocument.Clone(State.FilePath.analized, true)
          : sourceWordDocument;
      }
      catch (Exception ex)
      {
        MessageBox.Show(ex.Message, "Error");
      }

      if (document != null)
      {
        State.WDocument = document;

        WParse.Content();
        MessageBox.Show($"File {State.FilePath.withoutExtension} analysed", "Complete Status");
      }
    }

    private void RewriteCheckBox_OnClick(object sender, RoutedEventArgs e)
    {
      IsСomments = !IsСomments;
    }
  }
}