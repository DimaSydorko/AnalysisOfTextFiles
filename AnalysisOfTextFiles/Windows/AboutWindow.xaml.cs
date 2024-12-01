using System.Diagnostics;
using System.Windows;
using System.Windows.Navigation;

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