namespace TERO.ExcelHelper
{
  public class CellValueType
  {
    public string Value { get; set; }
    public CellType ValueCellType { get; set; }

    public enum CellType : int
    {
      Unknown = -1,
      Numeric = 0,
      String = 1,
      Date = 2,
      Blank = 3,
      Boolean = 4,
      Error = 5,
      Formula_Numeric = 6,
      Formula_String = 7,
      Formula_Boolean = 8,
      Formula_Date = 9,
      Formula_Error = 10
    }
  }
}