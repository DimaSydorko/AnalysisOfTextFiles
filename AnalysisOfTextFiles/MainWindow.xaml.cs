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
    AdminAuthWindow adminAuthWindow = new AdminAuthWindow();

    private Visibility _isAdminEditBtn, _isAdminAuthBtn, _isAdminChangePassBtn;

    public Visibility IsAdminEditBtn
    {
      get { return _isAdminEditBtn; }
      set
      {
        if (_isAdminEditBtn != value)
        {
          _isAdminEditBtn = value;
          OnPropertyChanged("IsAdminEditBtn");
        }
      }
    }

    public Visibility IsAdminAuthBtn
    {
      get { return _isAdminAuthBtn; }
      set
      {
        if (_isAdminAuthBtn != value)
        {
          _isAdminAuthBtn = value;
          OnPropertyChanged("IsAdminAuthBtn");
        }
      }
    }

    public Visibility IsAdminChangePassBtn
    {
      get { return _isAdminChangePassBtn; }
      set
      {
        if (_isAdminChangePassBtn != value)
        {
          _isAdminChangePassBtn = value;
          OnPropertyChanged("IsAdminChangePassBtn");
        }
      }
    }

    public MainWindow()
    {
      IsAdminAuthBtn = !State.IsAdminAuth && AdminSettings.IsUserAdmin() ? Visibility.Visible : Visibility.Collapsed;
      IsAdminEditBtn = State.IsAdminAuth ? Visibility.Visible : Visibility.Collapsed;
      IsAdminChangePassBtn = State.IsAdminAuth ? Visibility.Visible : Visibility.Collapsed;

      InitializeComponent();
      DataContext = this;

      adminAuthWindow.IsAdminAuthBtn += (visibility) =>
      {
        IsAdminAuthBtn = visibility ? Visibility.Visible : Visibility.Collapsed;
      };
      adminAuthWindow.IsAdminChangePassBtn += (visibility) =>
      {
        IsAdminChangePassBtn = visibility ? Visibility.Visible : Visibility.Collapsed;
      };
      adminAuthWindow.IsAdminEditBtn += (visibility) =>
      {
        IsAdminEditBtn = visibility ? Visibility.Visible : Visibility.Collapsed;
      };
    }

    protected void OnPropertyChanged(string propertyName)
    {
      PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
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

        string timeInfo = AdminSettings.IsUserAdmin() ? $" for {elapsedTime.TotalSeconds} s" : "";
        MessageBox.Show($"File {State.FilePath.withoutExtension} analysed{timeInfo}", "Complete Status");
      }
    }

    private void RewriteCheckBox_OnClick(object sender, RoutedEventArgs e)
    {
      IsСomments = !IsСomments;
    }

    private void AdminEdit_OnClick(object sender, RoutedEventArgs e)
    {
      EditorWindow modalWindow = new EditorWindow();
      modalWindow.Owner = this;
      modalWindow.ShowDialog();
    }

    private void AdminAuth_OnClick(object sender, RoutedEventArgs e)
    {
      adminAuthWindow.Owner = this;
      adminAuthWindow.ShowDialog();
    }

    private void AdminChangePass_OnClick(object sender, RoutedEventArgs e)
    {
      AdminChangePassWindow adminChangePass = new AdminChangePassWindow();
      adminChangePass.Owner = this;
      adminChangePass.ShowDialog();
    }
  }
}