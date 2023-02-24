using System.Collections.Generic;
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

  public static WStyle GetStyleFromEncoded(List<WStyle> allStyles, string encoded)
  {
    return allStyles.SingleOrDefault(s => { return s.encoded == encoded; });
  }
}
// public class StyleIssue
// {
//   public string styleId { get; set; }
//   public StyleName styleName { get; set; }
//   public bool isUsed { get; set; }
// }