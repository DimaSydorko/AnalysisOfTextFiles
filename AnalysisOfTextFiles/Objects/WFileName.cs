using System.IO;
using Microsoft.Win32;

namespace AnalysisOfTextFiles.Objects;

public class WFileName
{
  public string directory { get; set; }
  public string full { get; set; }
  public string extension { get; set; }
  public string withoutExtension { get; set; }
  public string analized { get; set; }

  public void Open()
  {
    OpenFileDialog openFileDialog = new OpenFileDialog();
    openFileDialog.Filter = "*.doc|*.docx";
    openFileDialog.InitialDirectory = @"c:\temp\";
    if (openFileDialog.ShowDialog() == true)
    {
      full = openFileDialog.FileName;
      directory = Path.GetDirectoryName(full);
      extension = Path.GetExtension(full);
      withoutExtension = Path.GetFileNameWithoutExtension(full);
      analized = $"{directory}/{withoutExtension} ANALYSED{extension}";
    }
  }
}