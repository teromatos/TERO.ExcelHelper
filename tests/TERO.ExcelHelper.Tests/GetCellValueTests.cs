using NUnit.Framework;
using TERO.Common;
using TERO.ExcelHelper;

namespace tests
{
  /// <summary>
  /// Summary description for GetCellTypeTests
  /// </summary>
  [TestFixture]
  public class GetCellValueTests : BaseTests
  {
    public GetCellValueTests()
    {
    }

    [Test]
    public void Test_GetCellValues()
    {
      ExcelWrapper.Open(TestFile);
      TestCellValue(1, 4, "True");
      TestCellValue(1, 5, "12345.12");
      TestCellValue(1, 6, "123456789");
      TestCellValue(1, 7, "1/1/2014 12:00:00 AM");
      TestCellValue(1, 8, "3/2/2014 12:00:00 AM");
      TestCellValue(1, 9, "False");
      TestCellValue(1, 10, "True");
      TestCellValue(1, 11, "test");
      TestCellValue(1, 12, "0123");
      TestCellValue(1, 13, "12345.12");
      TestCellValue(1, 14, "41640");
      TestCellValue(1, 15, "False");
      TestCellValue(1, 16, "test");
      TestCellValue(1, 17, "#ERROR!");
      TestCellValue(1, 18, "#ERROR!");

      TestCellValue(2, 4, "False");
      TestCellValue(2, 5, "12345.12");
      TestCellValue(2, 6, "123456789");
      TestCellValue(2, 7, "1/1/2014 12:00:00 AM");
      TestCellValue(2, 8, "3/2/2014 12:00:00 AM");
      TestCellValue(2, 9, "False");
      TestCellValue(2, 10, "True");
      TestCellValue(2, 11, "test");
      TestCellValue(2, 12, "0123");
      TestCellValue(2, 13, "123456789");
      TestCellValue(2, 14, "41700");
      TestCellValue(2, 15, "True");
      TestCellValue(2, 16, "0123");
      TestCellValue(2, 17, "#ERROR!");
      TestCellValue(2, 18, "#ERROR!");
    }

    private void TestCellValue(int row, int col, string cellValue)
    {
      var cellType = ExcelWrapper.GetCellValue(row, col);
      Assert.AreEqual(cellType, cellValue, "Row: {0}, Col: {1}".Fmt(row, col));
    }
  }
}