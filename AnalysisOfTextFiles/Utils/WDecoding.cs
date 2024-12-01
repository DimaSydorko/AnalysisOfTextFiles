using System.Collections.Generic;

public class WDecoding
{
  public static string? RemoveSuffixIfExists(string? input)
  {
    var suffix = " Char";
    var suffixUkr = " Знак";

    if (input != null && (input.EndsWith(suffix) || input.EndsWith(suffixUkr)))
      return input.Substring(0, input.Length - suffix.Length);

    return input;
  }


  public static string? GetOldDecStyle(string? encoded)
  {
    var entryTable = new Dictionary<string, string>
    {
      { "21", "TOC1" },
      { "22", "TOC2" },
      { "23", "TOC3" },
      { "13", "Normal" },
      { "1", "Heading1" },
      { "2", "Heading2" },
      { "3", "Heading3" }
    };

    if (encoded != null && entryTable.ContainsKey(encoded))
    {
      var entry = entryTable[encoded];
      return entry;
    }

    return null;
  }
}