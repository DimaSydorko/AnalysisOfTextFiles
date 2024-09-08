using System.Linq;

namespace AnalysisOfTextFiles.Objects;

public class WStyle
{
  public string decoded { get; set; }
  public string encoded { get; set; }

  public void SetDec(string dec)
  {
    decoded = dec;
  }

  public void SetEnc(string enc)
  {
    encoded = enc;
  }

  public static WStyle GetStyleFromEncoded(string encoded)
  {
    return State.Styles.SingleOrDefault(s => { return s.encoded == encoded; });
  }
}