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
}