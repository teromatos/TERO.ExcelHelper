using NUnit.Framework;
using TERO.Common;
using TERO.ExcelHelper;

namespace tests
{
  /// <summary>
  /// Summary description for GetCellTypeTests
  /// </summary>
  [TestFixture]
  public class GetCellTypeTests : BaseTests
  {
    public GetCellTypeTests()
    {
    }

    [Test]
    public void Test_GetCellType_Numeric()
    {
      ExcelWrapper.Open(TestFile);
      TestCellType(1, 5, CellValueType.CellType.Numeric);
      TestCellType(2, 5, CellValueType.CellType.Numeric);
      TestCellType(1, 6, CellValueType.CellType.Numeric);
      TestCellType(2, 6, CellValueType.CellType.Numeric);
    }

    [Test]
    public void Test_GetCellType_Date()
    {
      ExcelWrapper.Open(TestFile);
      TestCellType(1, 7, CellValueType.CellType.Date);
      TestCellType(2, 7, CellValueType.CellType.Date);
      TestCellType(1, 8, CellValueType.CellType.Date);
      TestCellType(2, 8, CellValueType.CellType.Date);
    }

    [Test]
    public void Test_GetCellType_String()
    {
      ExcelWrapper.Open(TestFile);
      TestCellType(1, 11, CellValueType.CellType.String);
      TestCellType(2, 11, CellValueType.CellType.String);
      TestCellType(1, 12, CellValueType.CellType.String);
      TestCellType(2, 12, CellValueType.CellType.String);
    }

    [Test]
    public void Test_GetCellType_Boolean()
    {
      ExcelWrapper.Open(TestFile);
      TestCellType(1, 9, CellValueType.CellType.Boolean);
      TestCellType(2, 9, CellValueType.CellType.Boolean);
    }

    [Test]
    public void Test_GetCellType_Formula_Boolean()
    {
      ExcelWrapper.Open(TestFile);
      TestCellType(1, 10, CellValueType.CellType.Formula_Boolean);
      TestCellType(2, 10, CellValueType.CellType.Formula_Boolean);
      TestCellType(1, 15, CellValueType.CellType.Formula_Boolean);
      TestCellType(2, 15, CellValueType.CellType.Formula_Boolean);
    }

    [Test]
    public void Test_GetCellType_Formula_Numeric()
    {
      ExcelWrapper.Open(TestFile);
      TestCellType(1, 13, CellValueType.CellType.Formula_Numeric);
      TestCellType(2, 13, CellValueType.CellType.Formula_Numeric);
      TestCellType(1, 14, CellValueType.CellType.Formula_Numeric);
      TestCellType(2, 14, CellValueType.CellType.Formula_Numeric);
    }

    [Test]
    public void Test_GetCellType_Formula_String()
    {
      ExcelWrapper.Open(TestFile);
      TestCellType(1, 16, CellValueType.CellType.Formula_String);
      TestCellType(2, 16, CellValueType.CellType.Formula_String);
    }

    [Test]
    public void Test_GetCellType_Formula_Error()
    {
      ExcelWrapper.Open(TestFile);
      TestCellType(1, 17, CellValueType.CellType.Formula_Error);
      TestCellType(2, 17, CellValueType.CellType.Formula_Error);
      TestCellType(1, 18, CellValueType.CellType.Formula_Error);
      TestCellType(2, 18, CellValueType.CellType.Formula_Error);
    }

    private void TestCellType(int row, int col, CellValueType.CellType exptectedCellType)
    {
      var cellType = ExcelWrapper.GetCellType(row, col);
      Assert.AreEqual(cellType, exptectedCellType, "Row: {0}, Col: {1}".Fmt(row, col));
    }
  }
}