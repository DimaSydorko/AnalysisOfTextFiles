using System.Collections.Generic;
using System.IO;
using Microsoft.Win32;

namespace AnalysisOfTextFiles.Objects;

public class WFilePath
{
  public string directory { get; set; }
  public string full { get; set; }
  public string extension { get; set; }
  public string withoutExtension { get; set; }
  public string analyzed { get; set; }
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
      path.analyzed = $"{path.directory}/{path.withoutExtension} ANALYSED.docx";
      path.report = $"{path.directory}/{path.withoutExtension} Report.txt";
      path.stylesSettings = $"{path.directory}/{path.withoutExtension} Styles Settings.txt";
    }

    return path;
  }
  
  public static List<WFilePath> OpenMultiple()
  {
    var paths = new List<WFilePath>();

    var openFileDialog = new OpenFileDialog
    {
      Filter = "Word Files (*.doc;*.docx)|*.doc;*.docx",
      Multiselect = true, // Enable multiple file selection
      InitialDirectory = @"c:\temp\"
    };

    if (openFileDialog.ShowDialog() == true)
    {
      foreach (var fileName in openFileDialog.FileNames)
      {
        var path = new WFilePath
        {
          full = fileName,
          directory = Path.GetDirectoryName(fileName),
          extension = Path.GetExtension(fileName),
          withoutExtension = Path.GetFileNameWithoutExtension(fileName),
          analyzed = $"{Path.GetDirectoryName(fileName)}/{Path.GetFileNameWithoutExtension(fileName)} ANALYSED.docx",
          report = $"{Path.GetDirectoryName(fileName)}/{Path.GetFileNameWithoutExtension(fileName)} Report.txt",
          stylesSettings = $"{Path.GetDirectoryName(fileName)}/{Path.GetFileNameWithoutExtension(fileName)} Styles Settings.txt"
        };
        paths.Add(path);
      }
    }

    return paths;
  }
}
