using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AnalysisOfTextFiles.Objects;

public class CheckEdited
{
  public static bool HasTextStyleChanged(Paragraph paragraph)
  {
    var runs = paragraph.Descendants<Run>().ToList();
    if (runs.Count <= 1) return false;
    var firstRunProps = runs[0].RunProperties;

    for (var i = 1; i < runs.Count; i++)
    {
      var currentRunProps = runs[i].RunProperties;
      if (!AreRunPropertiesEqual(firstRunProps, currentRunProps)) return true;
    }

    return false;
  }

  private static bool AreRunPropertiesEqual(RunProperties? firstProps, RunProperties? secondProps)
  {
    if (State.IsStrictMode)
    {
      if (firstProps == null && secondProps == null) return true;
      if (firstProps == null || secondProps == null) return false;
    }

    if (firstProps == null) return true;
    if (secondProps == null)
    {
      if (firstProps.Bold?.Val != null || firstProps.Italic?.Val != null || firstProps.FontSize?.Val != null ||
          firstProps.Color?.Val != null || firstProps.Underline?.Val != null)
        return false;
      return true;
    }

    if (firstProps.Bold?.Val != secondProps?.Bold?.Val) return false;
    if (firstProps.Italic?.Val != secondProps?.Italic?.Val) return false;
    if (firstProps.FontSize?.Val != secondProps?.FontSize?.Val) return false;
    if (firstProps.Color?.Val != secondProps?.Color?.Val) return false;
    if (firstProps.Underline?.Val != secondProps?.Underline?.Val) return false;

    return true;
  }

  public static bool IsEditedStyle(Paragraph paragraph)
  {
    var styleName = WDecoding.RemoveSuffixIfExists(CheckParagraph.GetParagraphStyle(paragraph));
    var isEdited = HasTextStyleChanged(paragraph);

    return styleName?.StartsWith(State.KeyWord) == true && isEdited;
  }
}