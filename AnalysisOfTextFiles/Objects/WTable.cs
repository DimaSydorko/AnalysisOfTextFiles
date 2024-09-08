namespace AnalysisOfTextFiles.Objects;

public class WTable
{
  public WTable(int idx, int rowIdx, int cellIdx, int parIdx)
  {
    Idx = idx;
    RowIdx = rowIdx;
    CellIdx = cellIdx;
    ParIdx = parIdx;
  }

  public int Idx { get; }
  public int RowIdx { get; }
  public int CellIdx { get; }
  public int ParIdx { get; }
}