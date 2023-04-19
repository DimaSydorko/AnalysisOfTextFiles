using System.Windows;

namespace AnalysisOfTextFiles;
public partial class EditorWindow
{
    public EditorWindow()
    {
        InitializeComponent();
        txtIniData.Text = AdminSettings.GetStyleSettings();
    }

    private void BtnSave_Click(object sender, RoutedEventArgs e)
    {
        string text = txtIniData.Text;
        AdminSettings.SetStyleSettings(text);

        DialogResult = true;
        Close();
    }

    private void BtnCancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}