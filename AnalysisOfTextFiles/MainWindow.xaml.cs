using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using AnalysisOfTextFiles.Objects;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

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
        // Open and clone file
        string copiedPath = $"{State.FilePath.directory}\\{State.FilePath.withoutExtension} CONVERTED.docx";;

        // Create a copy of the original document
        File.Copy(State.FilePath.full, copiedPath, true);

        // Create instance of OpenSettings
        OpenSettings openSettings = new OpenSettings();

        // Add the MarkupCompatibilityProcessSettings
        openSettings.MarkupCompatibilityProcessSettings =
          new MarkupCompatibilityProcessSettings(
            MarkupCompatibilityProcessMode.ProcessAllParts,
            FileFormatVersions.Office2016);
        
        // Open the copied document for modification
        using WordprocessingDocument copiedDoc = WordprocessingDocument.Open(copiedPath, true, openSettings);
        
        // Access the SettingsPart of the copied document
        DocumentSettingsPart? settingsPart = copiedDoc.MainDocumentPart.GetPartsOfType<DocumentSettingsPart>().FirstOrDefault();
        if (settingsPart != null)
        {
          // Create the Compatibility element
          Compatibility compatibility = new Compatibility(
            new CompatibilitySetting()
            {
              Name = new EnumValue<CompatSettingNameValues>(CompatSettingNameValues.CompatibilityMode),
              Uri = new StringValue("http://schemas.microsoft.com/office/word"),
              Val = new StringValue("16")
            });

          // Replace the existing Compatibility element or add it if it doesn't exist
          settingsPart.Settings.RemoveAllChildren<Compatibility>();
          settingsPart.Settings.AppendChild(compatibility);

          // Save the changes
          settingsPart.Settings.Save();
        }
        copiedDoc.Save();
        
        document = IsСomments
          ? (WordprocessingDocument)copiedDoc.Clone(State.FilePath.analized, true, openSettings)
          : copiedDoc;
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