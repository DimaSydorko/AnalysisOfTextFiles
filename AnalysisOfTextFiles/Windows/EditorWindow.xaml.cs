﻿using System.Diagnostics;
using System.Windows;
using System.Windows.Navigation;

namespace AnalysisOfTextFiles;

public partial class EditorWindow
{
  public EditorWindow()
  {
    InitializeComponent();
    var settingsData = AdminSettings.GetStyleData();
    txtIniData.Text = AdminSettings.GetStyleSettings(settingsData);
    keyWord.Text = AdminSettings.GetStyleKeyWord(settingsData);
  }

  private void BtnSave_Click(object sender, RoutedEventArgs e)
  {
    var text = txtIniData.Text;
    var keyWordText = keyWord.Text;
    AdminSettings.SetStyleSettings(text, keyWordText);

    DialogResult = true;
    Hide();
  }

  private void BtnCancel_Click(object sender, RoutedEventArgs e)
  {
    DialogResult = false;
    Hide();
  }

  private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
  {
    Process.Start(new ProcessStartInfo
    {
      FileName = e.Uri.AbsoluteUri,
      UseShellExecute = true
    });
    e.Handled = true;
  }
}