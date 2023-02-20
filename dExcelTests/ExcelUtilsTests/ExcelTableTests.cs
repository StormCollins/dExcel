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

    private readonly object[,] _discountFactorsWithoutHeadersTable =
    {
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
    public void GetTableLabel_Test()
    {
        Assert.AreEqual("Example Table", ExcelTableUtils.GetTableLabel(_parameterTable));
    }

    [Test]
    public void GetColumnHeaders_Test()
    {
        Assert.AreEqual(new List<string> {"PARAMETER", "VALUE"}, ExcelTableUtils.GetColumnHeaders(_parameterTable));
    }

    [Test]
    public void GetColumn_DateTimeTest()
    {
        Assert.AreEqual(
            expected: new List<DateTime> { new(2022, 06, 01), new(2022, 07, 01), new(2022, 08, 01) },
            actual: ExcelTableUtils.GetColumn<DateTime>(_discountFactorsTable, "Dates", 1));
    }
    
    [Test]
    public void GetColumn_DoubleTest()
    {
        Assert.AreEqual(
            expected: new List<double> { 1.000, 0.999, 0.998 },
            actual: ExcelTableUtils.GetColumn<double>(_discountFactorsTable, "Discount Factors", 1));
    }
    
    [Test]
    public void GetColumn_IntTest()
    {
        Assert.AreEqual(
            expected: new List<double> { 1, 2, 3 },
            actual: ExcelTableUtils.GetColumn<int>(_primeTable, "Position", 1));
    }

    [Test]
    public void GetColumn_StringTest()
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
            actual: ExcelTableUtils.GetColumn<string>(_parameterTable, "Parameter", 1));
    }

    [Test]
    public void GetColumn_ByColumnIndexTest()
    {
        List<string> expected =
            new()
            {
                "Example Table",
                "Parameter",
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
            };
        
        List<string>? actual = ExcelTableUtils.GetColumn<string>(_parameterTable, 0); 
        
        Assert.AreEqual(expected, actual);
    }

    [Test]
    public void GetColumn_ByColumnIndexAsDateTimeTest()
    {
        List<DateTime> expected =
            new()
            {
                DateTime.FromOADate(44713),
                DateTime.FromOADate(44743),
                DateTime.FromOADate(44774),
            };
        
        List<DateTime>? actual = ExcelTableUtils.GetColumn<DateTime>(_discountFactorsWithoutHeadersTable, 0); 
        Assert.AreEqual(expected, actual);
    }
    
    [Test]
    public void GetRowHeaders_Test()
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
            actual: ExcelTableUtils.GetRowHeaders(_parameterTable));
    }
    
    #region Lookup single value
    [Test]
    public void LookUpTableValueDateTest()
    {
        Assert.AreEqual(
            expected: 0.999,
            actual: ExcelTableUtils.GetTableValue<double>(_discountFactorsTable, "Discount Factors", "44743")); 
    }

    [Test]
    public void LookUpTableValueIntTest()
    {
        Assert.AreEqual(
            expected: 3,
            actual: ExcelTableUtils.GetTableValue<int>(_primeTable, "Primes", "2")); 
    }
    
    [Test]
    public void LookupTableValueStringTest()
    {
        Assert.AreEqual(
            expected: "LogLinear",
            actual: ExcelTableUtils.GetTableValue<string>(_parameterTable, "Value", "Interpolation")); 
    }
    
    [Test]
    public void LookUpNonExistentTableValueStringTest()
    {
        Assert.AreEqual(
            expected: null,
            actual: ExcelTableUtils.GetTableValue<string>(_parameterTable, "NotThere", "Interpolation")); 
    }

    [Test]
    public void LookUpBusinessDayConventionTest()
    {
        Assert.AreEqual(
            expected: BusinessDayConvention.ModifiedFollowing,
            actual: ExcelTableUtils.GetTableValue<BusinessDayConvention>(_parameterTable, "Value", "ValidBusinessDayConvention")); 
    }

    [Test]
    public void LookUpInvalidBusinessDayConventionTest()
    {
        Assert.Throws<ArgumentException>(() => 
            ExcelTableUtils.GetTableValue<BusinessDayConvention>(_parameterTable, "Value", "InvalidBusinessDayConvention"));
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
        return ExcelTableUtils.GetTableValue<DayCounter>(_parameterTable, "Value", label);
    }

    [Test]
    public void LookUpInvalidDayCountConventionTest()
    {
        Assert.Throws<ArgumentException>(() => 
            ExcelTableUtils.GetTableValue<DayCounter>(_parameterTable, "Value", "InvalidDayCountConvention"));
    }
    #endregion 

    #region Lookup multiple values
    [Test]
    public void LookUpMultiplyMappedTableValuesTest()
    {
        // Here we test one value in the 'Parameter' column mapping to multiple values in the 'Value' column.
        Assert.AreEqual(
            expected: new List<string> {"Deposits", "FRAs", "Interest Rate Swaps"},
            actual: ExcelTableUtils.LookUpTableValues<string>(_parameterTable, "Value", "Instruments"));
    }
    
    [Test]
    public void LookUpNonExistentTableValuesTest()
    {
        Assert.AreEqual(
            expected: null,
            actual: ExcelTableUtils.LookUpTableValues<string>(_parameterTable, "Value", "NotThere"));
    }

    [Test]
    public void LookUpNonExistentTableColumnHeaderTest()
    {
        Assert.AreEqual(
            expected: null,
            actual: ExcelTableUtils.LookUpTableValues<string>(_parameterTable, "NotThere", "Instruments"));
    }
    #endregion 
}
