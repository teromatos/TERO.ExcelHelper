using System;
using System.IO;
using NUnit.Framework;
using TERO.BaseManager;
using TERO.Common;
using TERO.ExcelHelper;

namespace tests
{
  /// <summary>
  /// Summary description for UnitTest1
  /// </summary>
  [TestFixture]
  public class BaseTests
  {
    public ExcelWrapper ExcelWrapper;
    public string OriginalTestFile = @"";
    public string TestFile = @"";
    public string SaveAsFile = @"";
    public string CreateFile = @"";

    public BaseTests()
    {
    }

    //Use TestInitialize to run code before running each test 
    [SetUp]
    public void MyTestInitialize()
    {
      OriginalTestFile = Path.Combine(Environment.CurrentDirectory, "test_original.xlsx");
      TestFile = Path.Combine(Environment.CurrentDirectory, "test.xlsx");
      SaveAsFile = Path.Combine(Environment.CurrentDirectory, "test_saveas.xlsx");
      CreateFile = Path.Combine(Environment.CurrentDirectory, "test_create.xlsx");

      ExcelWrapper = new ExcelWrapper();
      ExcelWrapper.Status += ExcelWrapper_Status;
      File.Copy(OriginalTestFile, TestFile, true);
    }

    private void ExcelWrapper_Status(object sender, StatusEventArgs e)
    {
      Console.WriteLine(e.Message);
    }

    //Use TestCleanup to run code after each test has run
    [TearDown]
    public void MyTestCleanup()
    {
      ExcelWrapper.Status -= ExcelWrapper_Status;
      ExcelWrapper = null;
    }
    
    ///
    /// 
    /// 
    /// 
    /// 
    /// 

    public void UpdateFile()
    {
      var row = ExcelWrapper.ExcelBook.GetSheetAt(0).GetRow(2);
      var cell = row.GetCell(1) ?? row.CreateCell(1);
      cell.SetCellValue("x");
    }
  }
}
