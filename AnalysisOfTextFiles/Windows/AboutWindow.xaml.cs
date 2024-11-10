using System.Windows;

namespace AnalysisOfTextFiles;

public partial class AboutWindow
{
  public AboutWindow()
  {
    InitializeComponent();
  }

  private void BtnClose_Click(object sender, RoutedEventArgs e)
  {
    Hide();
  }
  
  private void Hyperlink_RequestNavigate(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
  {
    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
    {
      FileName = e.Uri.AbsoluteUri,
      UseShellExecute = true
    });
    e.Handled = true;
  }
}
