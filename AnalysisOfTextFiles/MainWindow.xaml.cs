using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows;
using AnalysisOfTextFiles.Objects;
using DocumentFormat.OpenXml.Packaging;

namespace AnalysisOfTextFiles
{
  public partial class MainWindow : INotifyPropertyChanged
  {
    public event PropertyChangedEventHandler PropertyChanged;
    private static bool IsСomments { get; set; } = true;
    public Visibility IsAdmin { get; } = Visibility.Collapsed;
    public MainWindow()
    {
      IsAdmin = AdminSettings.IsUserAdmin() ? Visibility.Visible : Visibility.Collapsed;
      InitializeComponent();
      DataContext = this;
    }
    
    private void Upload_OnClick(object sender, RoutedEventArgs e)
    {
      State.FilePath = WFilePath.Open();
      State.IsСomments = IsСomments;

      if (State.FilePath.full == null) return;

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

        Stopwatch stopwatch = new Stopwatch();
        stopwatch.Start();
        
        WParse.Content();
        
        stopwatch.Stop();
        TimeSpan elapsedTime = stopwatch.Elapsed;
        
        MessageBox.Show($"File {State.FilePath.withoutExtension} analysed for {elapsedTime.TotalSeconds} s", "Complete Status");
      }
    }

    private void RewriteCheckBox_OnClick(object sender, RoutedEventArgs e)
    {
      IsСomments = !IsСomments;
    }

    private void AdminModal_OnClick(object sender, RoutedEventArgs e)
    {
      EditorWindow modalWindow = new EditorWindow();

      modalWindow.Owner = this;

      modalWindow.ShowDialog();
    }
  }
}