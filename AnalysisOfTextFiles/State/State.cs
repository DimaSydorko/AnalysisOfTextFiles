using System.Collections.Generic;
using AnalysisOfTextFiles.Objects;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

public class State
{
  public static bool IsСomments { get; set; } = true;
  public static bool IsAdminAuth { get; set; } = false;
  public static string KeyWord { get; set; } = "";
  public static string Content { get; set; } = "";
  public static Paragraph?  NextParagraph { get; set; } = null;
  public static Paragraph? PrevParagraph { get; set; } = null;
  public static List<WStyle> Styles { get; set; } = new();
  public static List<StyleProperties> StylesSettings { get; set; } = new();
  public static WFilePath FilePath { get; set; } = new();
  public static WordprocessingDocument WDocument { get; set; } = null;
}