using System.Windows;

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
    Close();
  }

  private void BtnCancel_Click(object sender, RoutedEventArgs e)
  {
    DialogResult = false;
    Close();
  }
}