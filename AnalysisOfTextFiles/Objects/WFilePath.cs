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

  public static WFilePath Open()
  {
    WFilePath path = new WFilePath();
    
    OpenFileDialog openFileDialog = new OpenFileDialog();
    openFileDialog.Filter = "*.doc|*.docx";
    openFileDialog.InitialDirectory = @"c:\temp\";
    
    if (openFileDialog.ShowDialog() == true)
    {
      path.full = openFileDialog.FileName;
      path.directory = Path.GetDirectoryName(path.full);
      path.extension = Path.GetExtension(path.full);
      path.withoutExtension = Path.GetFileNameWithoutExtension(path.full);
      path.analized = $"{path.directory}/{path.withoutExtension} ANALYSED{path.extension}";
      path.report = $"{path.directory}/{path.withoutExtension} Report.txt";
    }

    return path;
  }
}