using NUnit.Framework;
using TERO.Common;

namespace tests
{
  /// <summary>
  /// Summary description for InstanceTests
  /// </summary>
  [TestFixture]
  public class InstanceTests : BaseTests
  {
    public InstanceTests()
    {
    }

    [Test]
    public void NewInstance_DoesNotContain_Null_OpenedExcelFilename()
    {
      Assert.IsNotNull(ExcelWrapper.OpenedExcelFilename);
    }

    [Test]
    public void NewInstance_Contains_Empty_OpenedExcelFilename()
    {
      Assert.IsTrue(ExcelWrapper.OpenedExcelFilename.OrdinalEquals(""));
    }

    [Test]
    public void NewInstance_Contains_LastGetValueIsNull_Equals_True()
    {
      Assert.IsTrue(ExcelWrapper.LastGetValueIsNull);
    }
  }
}