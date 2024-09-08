using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows;
using AnalysisOfTextFiles.Objects;
using DocumentFormat.OpenXml.Packaging;

namespace AnalysisOfTextFiles;

public partial class MainWindow : INotifyPropertyChanged
{
  private Visibility _isAdminEditBtn, _isAdminAuthBtn, _isAdminChangePassBtn;
  private readonly AdminAuthWindow adminAuthWindow = new();

  public MainWindow()
  {
    IsAdminAuthBtn = !State.IsAdminAuth && AdminSettings.IsUserAdmin() ? Visibility.Visible : Visibility.Collapsed;
    IsAdminEditBtn = State.IsAdminAuth ? Visibility.Visible : Visibility.Collapsed;
    IsAdminChangePassBtn = State.IsAdminAuth ? Visibility.Visible : Visibility.Collapsed;

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
  }

  private static bool IsСomments { get; set; } = true;

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

  public event PropertyChangedEventHandler PropertyChanged;

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
      using var sourceWordDocument = WordprocessingDocument.Open(State.FilePath.full, false);
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

      var stopwatch = new Stopwatch();
      stopwatch.Start();

      WParse.Content();

      stopwatch.Stop();
      var elapsedTime = stopwatch.Elapsed;

      var timeInfo = AdminSettings.IsUserAdmin() ? $" for {elapsedTime.TotalSeconds} s" : "";
      MessageBox.Show($"File {State.FilePath.withoutExtension} analysed{timeInfo}", "Complete Status");
    }
  }

  private void RewriteCheckBox_OnClick(object sender, RoutedEventArgs e)
  {
    IsСomments = !IsСomments;
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
}