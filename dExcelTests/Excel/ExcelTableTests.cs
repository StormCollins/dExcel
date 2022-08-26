﻿namespace dExcelTests.Excel;

using dExcel.ExcelUtils;
using NUnit.Framework;

[TestFixture]
public class ExcelTableTests
{
    private readonly object[,] _parameterTable =
    {
        { "Example Table", "" },
        { "Parameter", "Value" },
        { "Curve Name", "SingleCurve" },
        { "Interpolation", "LogLinear" },
        { "Instruments", "Deposits" },
        { "", "FRAs" },
        { "", "Interest Rate Swaps" },
        { "Base Date", "2022-06-01" },
    };

    private readonly object[,] _discountFactorsTable =
    {
        { "Discount Factors Table", "" },
        { "Dates", "Discount Factors" },
        { DateTime.FromOADate(44713), 1.000 },
        { DateTime.FromOADate(44743), 0.999 },
        { DateTime.FromOADate(44774), 0.998 },
    };

    private readonly object[,] _primeTable =
    {
        { "Prime Numbers", "" },
        { "Position", "Primes" },
        { 1, 2 },
        { 2, 3 },
        { 3, 5 },
    };
    
    [Test]
    public void GetTableLabelTest()
    {
        Assert.AreEqual("Example Table", ExcelTable.GetTableLabel(_parameterTable));
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
            actual: ExcelTable.GetColumn<int>(_primeTable, "Position", 1));
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
    
    // ------------------------------------------------------------------------------
    // LookUp Single Value
    [Test]
    public void LookUpTableValueDateTest()
    {
        Assert.AreEqual(
            expected: 0.999,
            actual: ExcelTable.LookUpTableValue<double>(_discountFactorsTable, "Discount Factors", DateTime.FromOADate(44743).ToString())); 
    }

    [Test]
    public void LookUpTableValueIntTest()
    {
        Assert.AreEqual(
            expected: 3,
            actual: ExcelTable.LookUpTableValue<int>(_primeTable, "Primes", "2")); 
    }
    
    [Test]
    public void LookupTableValueStringTest()
    {
        Assert.AreEqual(
            expected: "LogLinear",
            actual: ExcelTable.LookUpTableValue<string>(_parameterTable, "Value", "Interpolation")); 
    }
    
    [Test]
    public void LookUpNonExistentTableValueStringTest()
    {
        Assert.AreEqual(
            expected: null,
            actual: ExcelTable.LookUpTableValue<string>(_parameterTable, "NotThere", "Interpolation")); 
    }

    // --------------------------------------------------------------------------
    // LookUp Values
    [Test]
    public void LookUpMultiplyMappedTableValuesTest()
    {
        // Here we test one value in the 'Parameter' column mapping to multiple values in the 'Value' column.
        Assert.AreEqual(
            expected: new List<string> {"Deposits", "FRAs", "Interest Rate Swaps"},
            actual: ExcelTable.LookUpTableValues<string>(_parameterTable, "Value", "Instruments"));
    }
    
    [Test]
    public void LookUpNonExistentTableValuesTest()
    {
        Assert.AreEqual(
            expected: null,
            actual: ExcelTable.LookUpTableValues<string>(_parameterTable, "Value", "NotThere"));
    }

    [Test]
    public void LookUpNonExistentTableColumnHeaderTest()
    {
        Assert.AreEqual(
            expected: null,
            actual: ExcelTable.LookUpTableValues<string>(_parameterTable, "NotThere", "Instruments"));
    }
}