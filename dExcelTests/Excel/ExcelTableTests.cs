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
        {"Instruments", "Deposits"},
        {"", "FRAs"},
        {"", "Interest Rate Swaps"},
        {"Base Date", "2022-06-01"},
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
    public void GetColumnDateTest()
    {
        Assert.AreEqual(
            expected: new List<DateTime> { new(2022, 06, 01), new(2022, 07, 01), new(2022, 08, 01) },
            actual: ExcelTable.GetColumn<DateTime>(_discountFactorsTable, "Dates", 1));
    }
    
    [Test]
    public void GetColumnDoubleTest()
    {
        Assert.AreEqual(
            expected: new List<double> { 1.000, 0.999, 0.998 },
            actual: ExcelTable.GetColumn<double>(_discountFactorsTable, "Discount Factors", 1));
    }
    
    [Test]
    public void GetColumnIntTest()
    {
        Assert.AreEqual(
            expected: new List<double> { 1, 2, 3 },
            actual: ExcelTable.GetColumn<int>(_integerTable, "Int1", 1));
    }

    [Test]
    public void GetColumnStringTest()
    {
        Assert.AreEqual(
            expected: new List<string> {"Curve Name", "Interpolation", "Instruments", "", "", "Base Date"},
            actual: ExcelTable.GetColumn<string>(_parameterTable, "Parameter", 1));
    }

    [Test]
    public void GetRowHeadersTest()
    {
        Assert.AreEqual(
            expected: new List<string> {"Curve Name", "Interpolation", "Instruments", "", "", "Base Date"}, 
            actual: ExcelTable.GetRowHeaders(_parameterTable));
    }

    [Test]
    public void LookupTableValueIntTest()
    {
        Assert.AreEqual(
            expected: 3,
            actual: ExcelTable.LookupTableValue<int>(_integerTable, "Int2", "2")); 
    }
    
    [Test]
    public void LookupTableValueStringTest()
    {
        Assert.AreEqual(
            expected: "LogLinear",
            actual: ExcelTable.LookupTableValue<string>(_parameterTable, "Value", "Interpolation")); 
    }
    
    [Test]
    public void LookupNonExistentTableValueStringTest()
    {
        Assert.AreEqual(
            expected: null,
            actual: ExcelTable.LookupTableValue<string>(_parameterTable, "NotThere", "Interpolation")); 
    }

    [Test]
    public void LookupMultiplyMappedTableValuesTest()
    {
        // TODO: Explain this test.
        Assert.AreEqual(
            expected: new List<string> {"Deposits", "FRAs", "Interest Rate Swaps"},
            actual: ExcelTable.LookupTableValues<string>(_parameterTable, "Value", "Instruments"));
    }
    
    [Test]
    public void LookupNonExistentTableValuesTest()
    {
        Assert.AreEqual(
            expected: null,
            actual: ExcelTable.LookupTableValues<string>(_parameterTable, "Value", "NotThere"));
    }
}
