using System.Collections.Generic;
using AnalysisOfTextFiles.Objects;
using DocumentFormat.OpenXml.Packaging;

public class State
{
  public static bool IsСomments { get; set; } = true;
  public static bool IsStrictMode { get; set; } = false;
  public static bool IsAllowEmptyLine { get; set; } = false;
  public static bool IsAdminAuth { get; set; } = false;
  public static string KeyWord { get; set; } = "";
  public static string Content { get; set; } = "";
  public static List<WStyle> Styles { get; set; } = new();
  public static WPage PageSettings { get; set; } = null;
  public static string? NextParagraphName { get; set; } = null;
  public static string? PrevParagraphName { get; set; } = null;
  public static List<StyleProperties> StylesSettings { get; set; } = new();
  public static WFilePath FilePath { get; set; } = new();
  public static WordprocessingDocument WDocument { get; set; } = null;
}