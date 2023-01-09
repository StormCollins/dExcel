namespace dExcelTests.ExcelUtilsTests;

using dExcel.ExcelUtils;
using NUnit.Framework;
using QLNet;

[TestFixture]
public class ExcelTableTests
{
    private readonly object[,] _parameterTable =
    {
        { "Example Table", "" },
        { "Parameter", "Value" },
        { "CurveUtils Name", "SingleCurve" },
        { "Interpolation", "LogLinear" },
        { "Instruments", "Deposits" },
        { "", "FRAs" },
        { "", "Interest Rate Swaps" },
        { "Base Date", "2022-06-01" },
        { "ValidBusinessDayConvention", "ModifiedFollowing"},
        { "InvalidBusinessDayConvention", "Invalid"},
        { "Bus252DayCountConvention", "Bus252"},
        { "Act360DayCountConvention", "Act360"},
        { "Act364DayCountConvention", "Act364"},
        { "Act365DayCountConvention", "Act365"},
        { "ActActDayCountConvention", "ActAct"},
        { "InvalidDayCountConvention", "Invalid"},
    };

    private readonly object[,] _discountFactorsTable =
    {
        { "Discount Factors Table", "" },
        { "Dates", "Discount Factors" },
        { 44713, 1.000 },
        { 44743, 0.999 },
        { 44774, 0.998 },
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
        Assert.AreEqual(new List<string> {"PARAMETER", "VALUE"}, ExcelTable.GetColumnHeaders(_parameterTable));
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
            expected: new List<string>
            {
                "CurveUtils Name", 
                "Interpolation", 
                "Instruments", 
                "", 
                "", 
                "Base Date", 
                "ValidBusinessDayConvention", 
                "InvalidBusinessDayConvention",
                "Bus252DayCountConvention", 
                "Act360DayCountConvention", 
                "Act364DayCountConvention",
                "Act365DayCountConvention", 
                "ActActDayCountConvention", 
                "InvalidDayCountConvention",
            },
            actual: ExcelTable.GetColumn<string>(_parameterTable, "Parameter", 1));
    }

    [Test]
    public void GetRowHeadersTest()
    {
        Assert.AreEqual(
            expected: new List<string>
            {
                "CURVEUTILS NAME", 
                "INTERPOLATION", 
                "INSTRUMENTS", 
                "", 
                "", 
                "BASE DATE", 
                "VALIDBUSINESSDAYCONVENTION", 
                "INVALIDBUSINESSDAYCONVENTION",
                "BUS252DAYCOUNTCONVENTION", 
                "ACT360DAYCOUNTCONVENTION",
                "ACT364DAYCOUNTCONVENTION",
                "ACT365DAYCOUNTCONVENTION",
                "ACTACTDAYCOUNTCONVENTION",
                "INVALIDDAYCOUNTCONVENTION",
            }, 
            actual: ExcelTable.GetRowHeaders(_parameterTable));
    }
    
    #region Lookup single value
    [Test]
    public void LookUpTableValueDateTest()
    {
        Assert.AreEqual(
            expected: 0.999,
            actual: ExcelTable.GetTableValue<double>(_discountFactorsTable, "Discount Factors", "44743")); 
    }

    [Test]
    public void LookUpTableValueIntTest()
    {
        Assert.AreEqual(
            expected: 3,
            actual: ExcelTable.GetTableValue<int>(_primeTable, "Primes", "2")); 
    }
    
    [Test]
    public void LookupTableValueStringTest()
    {
        Assert.AreEqual(
            expected: "LogLinear",
            actual: ExcelTable.GetTableValue<string>(_parameterTable, "Value", "Interpolation")); 
    }
    
    [Test]
    public void LookUpNonExistentTableValueStringTest()
    {
        Assert.AreEqual(
            expected: null,
            actual: ExcelTable.GetTableValue<string>(_parameterTable, "NotThere", "Interpolation")); 
    }

    [Test]
    public void LookUpBusinessDayConventionTest()
    {
        Assert.AreEqual(
            expected: BusinessDayConvention.ModifiedFollowing,
            actual: ExcelTable.GetTableValue<BusinessDayConvention>(_parameterTable, "Value", "ValidBusinessDayConvention")); 
    }

    [Test]
    public void LookUpInvalidBusinessDayConventionTest()
    {
        Assert.Throws<ArgumentException>(() => 
            ExcelTable.GetTableValue<BusinessDayConvention>(_parameterTable, "Value", "InvalidBusinessDayConvention"));
    }

    public static IEnumerable<TestCaseData> LookUpDayCountConventionTestCaseData()
    {
        yield return new TestCaseData("Bus252DayCountConvention").Returns(new Business252());
        yield return new TestCaseData("Act360DayCountConvention").Returns(new Actual360());
        yield return new TestCaseData("Act365DayCountConvention").Returns(new Actual365Fixed());
        yield return new TestCaseData("ActActDayCountConvention").Returns(new ActualActual());
    }

    [Test]
    [TestCaseSource(nameof(LookUpDayCountConventionTestCaseData))]
    public DayCounter? LookUpDayCountConventionTest(string label)
    {
        return ExcelTable.GetTableValue<DayCounter>(_parameterTable, "Value", label);
    }

    [Test]
    public void LookUpInvalidDayCountConventionTest()
    {
        Assert.Throws<ArgumentException>(() => 
            ExcelTable.GetTableValue<DayCounter>(_parameterTable, "Value", "InvalidDayCountConvention"));
    }
    #endregion 

    #region Lookup multiple values
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
    #endregion 
}
