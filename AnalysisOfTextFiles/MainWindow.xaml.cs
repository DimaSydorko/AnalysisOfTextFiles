using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows;
using AnalysisOfTextFiles.Objects;
using DocumentFormat.OpenXml.Packaging;

namespace AnalysisOfTextFiles;

public partial class MainWindow : INotifyPropertyChanged
{
  private readonly AboutWindow aboutWindow = new();
  private readonly AdminAuthWindow adminAuthWindow = new();
  private Visibility _isAdminEditBtn, _isAdminAuthBtn, _isAdminChangePassBtn, _IsAdminGetDocSttingsBtn;

  public MainWindow()
  {
    IsAdminAuthBtn = !State.IsAdminAuth && AdminSettings.IsUserAdmin() ? Visibility.Visible : Visibility.Collapsed;
    IsAdminEditBtn = State.IsAdminAuth ? Visibility.Visible : Visibility.Collapsed;
    IsAdminChangePassBtn = State.IsAdminAuth ? Visibility.Visible : Visibility.Collapsed;
    IsAdminGetDocSttingsBtn = State.IsAdminAuth ? Visibility.Visible : Visibility.Collapsed;

    InitializeComponent();
    DataContext = this;

    adminAuthWindow.IsAdminAuthBtn += visibility =>
    {
      IsAdminAuthBtn = visibility ? Visibility.Visible : Visibility.Collapsed;
    };
    adminAuthWindow.IsAdminChangePassBtn += visibility =>
    {
      IsAdminChangePassBtn = visibility ? Visibility.Visible : Visibility.Collapsed;
    };
    adminAuthWindow.IsAdminEditBtn += visibility =>
    {
      IsAdminEditBtn = visibility ? Visibility.Visible : Visibility.Collapsed;
    };
    adminAuthWindow.IsAdminGetDocSttingsBtn += visibility =>
    {
      IsAdminGetDocSttingsBtn = visibility ? Visibility.Visible : Visibility.Collapsed;
    };
  }

  private static bool IsСomments { get; set; } = true;
  private static bool IsStrictMode { get; set; }
  private static bool IsAllowEmptyLine { get; set; }

  public Visibility IsAdminEditBtn
  {
    get => _isAdminEditBtn;
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
    get => _isAdminAuthBtn;
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
    get => _isAdminChangePassBtn;
    set
    {
      if (_isAdminChangePassBtn != value)
      {
        _isAdminChangePassBtn = value;
        OnPropertyChanged("IsAdminChangePassBtn");
      }
    }
  }

  public Visibility IsAdminGetDocSttingsBtn
  {
    get => _IsAdminGetDocSttingsBtn;
    set
    {
      if (_IsAdminGetDocSttingsBtn != value)
      {
        _IsAdminGetDocSttingsBtn = value;
        OnPropertyChanged("IsAdminGetDocSttingsBtn");
      }
    }
  }

  public event PropertyChangedEventHandler PropertyChanged;

  protected void OnPropertyChanged(string propertyName)
  {
    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
  }

  private async void Upload_OnClick(object sender, RoutedEventArgs e)
  {
    var filePaths = WFilePath.OpenMultiple();
    if (filePaths?.Count == 0) return;

    State.IsСomments = IsСomments;
    State.IsStrictMode = IsStrictMode;
    State.IsAllowEmptyLine = IsAllowEmptyLine;

    foreach (var filePath in filePaths)
    {
      State.FilePath = filePath;

      WordprocessingDocument document;
      try
      {
        using var sourceWordDocument = WordprocessingDocument.Open(State.FilePath.full, false);
        document = IsСomments
          ? (WordprocessingDocument)sourceWordDocument.Clone(State.FilePath.analyzed, true)
          : sourceWordDocument;
      }
      catch (Exception ex)
      {
        MessageBox.Show(ex.Message, "Error");
        continue;
      }

      State.WDocument = document;
      var stopwatch = new Stopwatch();
      stopwatch.Start();
      await Task.Run(() => WParse.Content());
      stopwatch.Stop();
      var elapsedTime = stopwatch.Elapsed;
      var timeInfo = AdminSettings.IsUserAdmin() ? $" for {elapsedTime.TotalSeconds} s" : "";
      MessageBox.Show($"File {State.FilePath.withoutExtension} analysed{timeInfo}", "Complete Status");
    }
  }

  public void GetStyles_OnClick(object sender, RoutedEventArgs e)
  {
    State.FilePath = WFilePath.Open();

    if (State.FilePath.full == null) return;

    WordprocessingDocument document = null;
    try
    {
      document = WordprocessingDocument.Open(State.FilePath.full, false);
    }
    catch (Exception ex)
    {
      MessageBox.Show(ex.Message, "Error");
    }

    if (document != null)
    {
      State.WDocument = document;

      var stopwatch = new Stopwatch();
      stopwatch.Start();

      WParse.StyleSettings();

      stopwatch.Stop();
      var elapsedTime = stopwatch.Elapsed;

      var timeInfo = AdminSettings.IsUserAdmin() ? $" for {elapsedTime.TotalSeconds} s" : "";
      MessageBox.Show($"Styles in file {State.FilePath.withoutExtension} analysed{timeInfo}", "Complete Status");
    }
  }

  private void RewriteCheckBox_OnClick(object sender, RoutedEventArgs e)
  {
    IsСomments = !IsСomments;
  }

  private void StrictCheckBox_OnClick(object sender, RoutedEventArgs e)
  {
    IsStrictMode = !IsStrictMode;
  } 
  
  private void EmptyLineCheckBox_OnClick(object sender, RoutedEventArgs e)
  {
    IsAllowEmptyLine = !IsAllowEmptyLine;
  }

  private void AdminEdit_OnClick(object sender, RoutedEventArgs e)
  {
    var modalWindow = new EditorWindow();
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
    var adminChangePass = new AdminChangePassWindow();
    adminChangePass.Owner = this;
    adminChangePass.ShowDialog();
  }

  private void About_OnClick(object sender, RoutedEventArgs e)
  {
    aboutWindow.Owner = this;
    aboutWindow.ShowDialog();
  }
}