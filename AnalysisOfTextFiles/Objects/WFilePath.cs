using System.IO;
using Microsoft.Win32;

namespace AnalysisOfTextFiles.Objects;

public class WFilePath
{
  public string directory { get; set; }
  public string full { get; set; }
  public string extension { get; set; }
  public string withoutExtension { get; set; }
  public string analized { get; set; }
  public string report { get; set; }
  public string stylesSettings { get; set; }

  public static WFilePath Open()
  {
    var path = new WFilePath();

    var openFileDialog = new OpenFileDialog();
    openFileDialog.Filter = "*.doc|*.docx";
    // openFileDialog.Filter = @"All Files|*.docx;*.doc;|Word File (.docx ,.doc)|*.docx;*.doc";
    openFileDialog.InitialDirectory = @"c:\temp\";

    if (openFileDialog.ShowDialog() == true)
    {
      path.full = openFileDialog.FileName;
      path.directory = Path.GetDirectoryName(path.full);
      path.extension = Path.GetExtension(path.full);
      path.withoutExtension = Path.GetFileNameWithoutExtension(path.full);
      path.analized = $"{path.directory}/{path.withoutExtension} ANALYSED.docx";
      path.report = $"{path.directory}/{path.withoutExtension} Report.txt";
      path.stylesSettings = $"{path.directory}/{path.withoutExtension} Styles Settings.txt";
    }

    return path;
  }
}