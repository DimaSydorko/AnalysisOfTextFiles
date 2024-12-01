namespace AnalysisOfTextFiles.Objects;

public class WStyle
{
  public string Decoded { get; set; }
  public string Encoded { get; set; }

  public void SetDec(string decoded)
  {
    Decoded = decoded;
  }

  public void SetEnc(string encoded)
  {
    Encoded = encoded;
  }

  private static string RemoveLastZero(string encoded)
  {
    if (!string.IsNullOrEmpty(encoded) && encoded.EndsWith("0")) return encoded.Substring(0, encoded.Length - 1);
    return encoded;
  }

  public static WStyle GetStyleFromEncoded(string encoded)
  {
    return State.Styles.Find(s => { return RemoveLastZero(s.Encoded) == encoded || s.Encoded == encoded; });
  }

  public static string GetDecodedStyle(string encoded)
  {
    return State.Styles.Find(s => { return RemoveLastZero(s.Encoded) == encoded || s.Encoded == encoded; })?.Decoded;
  }
}