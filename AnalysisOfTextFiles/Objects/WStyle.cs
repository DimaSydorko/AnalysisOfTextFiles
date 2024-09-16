using System.Linq;

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
  public static WStyle GetStyleFromEncoded(string encoded)
  {
    return State.Styles.SingleOrDefault(s => { return s.Encoded == encoded; });
  }  
  
  public static string GetDecodedStyle(string encoded)
  {
    return State.Styles.Find(s => { return s.Encoded == encoded; })?.Decoded;
  }
}