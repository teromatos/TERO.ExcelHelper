using System;
using System.IO;
using NUnit.Framework;
using TERO.Common;

namespace tests
{
  /// <summary>
  /// Summary description for OpenCloseSaveTests
  /// </summary>
  [TestFixture]
  public class OpenCloseSaveTests : BaseTests
  {
    public OpenCloseSaveTests()
    {
    }
    
    [Test]
    public void test_open_close()
    {
      var currFileInfo = new FileInfo(TestFile);
      var currLen = currFileInfo.Length;
      var currDate = currFileInfo.LastWriteTime;
      var currBytes = File.ReadAllBytes(TestFile);

      ExcelWrapper.Open(TestFile);
      UpdateFile();
      ExcelWrapper.Close();

      var newFileInfo = new FileInfo(TestFile);
      var newLen = currFileInfo.Length;
      var newDate = currFileInfo.LastWriteTime;
      var newBytes = File.ReadAllBytes(TestFile);

      Assert.AreEqual(currLen, newLen);
      Assert.AreEqual(currDate, newDate);
      Assert.AreEqual(currBytes.Length, newBytes.Length);
      Console.WriteLine("Current: {0}, New: {1}", currBytes.Length, newBytes.Length);
    }

    [Test]
    public void test_open_save()
    {
      var currBytes = File.ReadAllBytes(TestFile);
      
      ExcelWrapper.Open(TestFile);
      UpdateFile();
      ExcelWrapper.Save();

      var newBytes = File.ReadAllBytes(TestFile);

      Assert.AreNotEqual(currBytes.Length, newBytes.Length);
      Console.WriteLine("Current: {0}, New: {1}", currBytes.Length, newBytes.Length);
    }

    [Test]
    public void test_open_saveas()
    {
      var currBytes = File.ReadAllBytes(TestFile);

      File.Delete(SaveAsFile);

      ExcelWrapper.Open(TestFile);
      UpdateFile();
      ExcelWrapper.SaveAs(SaveAsFile);

      var newBytes = File.ReadAllBytes(TestFile);
      Assert.AreEqual(currBytes.Length, newBytes.Length);
      Assert.IsTrue(File.Exists(SaveAsFile));
      Assert.IsTrue(ExcelWrapper.OpenedExcelFilename.OrdinalEquals(SaveAsFile));

      Console.WriteLine("Current: {0}, New: {1}", currBytes.Length, newBytes.Length);

      File.Delete(SaveAsFile);
    }

    [Test]
    [ExpectedException(typeof(ApplicationException))]
    public void test_create_with_existing_file_throws_exception()
    {
      ExcelWrapper.Create(TestFile);
      Assert.IsTrue(File.Exists(TestFile));
    }

    [Test]
    public void test_create_only_will_not_create_file()
    {
      File.Delete(CreateFile);
      ExcelWrapper.Create(CreateFile);
      Assert.IsTrue(!File.Exists(CreateFile));
    }

    [Test]
    public void test_create_and_close_will_not_create_file()
    {
      File.Delete(CreateFile);
      ExcelWrapper.Create(CreateFile);
      ExcelWrapper.Close();
      Assert.IsTrue(!File.Exists(CreateFile));
    }

    [Test]
    public void test_create_and_save_will_create_file()
    {
      File.Delete(CreateFile);
      ExcelWrapper.Create(CreateFile);
      ExcelWrapper.Save();
      Assert.IsTrue(File.Exists(CreateFile));
    }

    [Test]
    public void test_create_and_saveas_will_create_file()
    {
      File.Delete(CreateFile);
      ExcelWrapper.Create(CreateFile);
      ExcelWrapper.SaveAs(SaveAsFile);
      Assert.IsTrue(!File.Exists(CreateFile));
      Assert.IsTrue(File.Exists(SaveAsFile));
    }

  }
}
