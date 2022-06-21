namespace dExcelTests.Excel;

using dExcel.ExcelUtils;
using NUnit.Framework;

[TestFixture]
public class ExcelTableTests
{
    private readonly object[,] _parameterTable =
    {
        {"Example Table", ""},
        {"Parameter", "Value"},
        {"Curve Name", "SingleCurve"},
        {"Interpolation", "LogLinear"},
    };

    private readonly object[,] _discountFactorsTable =
    {
        {"Discount Factors Table", ""},
        {"Dates", "Discount Factors"},
        {44713, "1.000"},
        {44743, "0.999"},
        {44774, "0.998"},
    };

    private readonly object[,] _integerTable =
    {
        {"Integer Table", ""},
        {"Int1", "Int2"},
        {"1", "2"},
        {"2", "3"},
        {"3", "4"},
    };
    
    [Test]
    public void GetTableTypeTest()
    {
        Assert.AreEqual("Example Table", ExcelTable.GetTableType(_parameterTable));
    }

    [Test]
    public void GetColumnTitlesTest()
    {
        Assert.AreEqual(new List<string> {"Parameter", "Value"}, ExcelTable.GetColumnHeaders(_parameterTable));
    }

    [Test]
    public void GetStringColumnTest()
    {
        Assert.AreEqual(
            expected: new List<string> {"Curve Name", "Interpolation"},
            actual: ExcelTable.GetColumn<string>(_parameterTable, "Parameter"));
    }

    [Test]
    public void GetDoubleColumnTest()
    {
        Assert.AreEqual(
            expected: new List<double> { 1.000, 0.999, 0.998 },
            actual: ExcelTable.GetColumn<double>(_discountFactorsTable, "Discount Factors"));
    }
    
    [Test]
    public void GetDateColumnTest()
    {
        Assert.AreEqual(
            expected: new List<DateTime> { new(2022, 06, 01), new(2022, 07, 01), new(2022, 08, 01) },
            actual: ExcelTable.GetColumn<DateTime>(_discountFactorsTable, "Dates"));
    }
    
    [Test]
    public void GetIntColumnTest()
    {
        Assert.AreEqual(
            expected: new List<double> { 1, 2, 3 },
            actual: ExcelTable.GetColumn<int>(_integerTable, "Int1"));
    }
}
