using System.Collections.Generic;
using AnalysisOfTextFiles.Objects;
using DocumentFormat.OpenXml.Packaging;

public class State
{
  public static bool IsСomments { get; set; } = true;
  public static bool IsAdminAuth { get; set; } = false;
  public static string KeyWord { get; set; } = "";
  public static string Content { get; set; } = "";
  public static List<WStyle> Styles { get; set; } = new();
  public static WFilePath FilePath { get; set; } = new();
  public static WordprocessingDocument WDocument { get; set; } = null;
}