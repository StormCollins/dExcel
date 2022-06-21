namespace dExcelTests.Excel;

using dExcel.ExcelUtils;
using NUnit.Framework;

[TestFixture]
public class ExcelTableTests
{
    private readonly object[,] _exampleTable =
    {
        {"Example Table", ""},
        {"Parameter", "Value"},
    };
    
    [Test]
    public void GetTableTypeTest()
    {
        Assert.AreEqual("Example Table", ExcelTable.GetTableType(_exampleTable));
    }

    [Test]
    public void GetColumnTitlesTest()
    {
        Assert.AreEqual(new List<string> {"Parameter", "Value"}, ExcelTable.GetColumnHeaders(_exampleTable));
    }
}
