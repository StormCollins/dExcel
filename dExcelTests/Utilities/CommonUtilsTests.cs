namespace dExcelTests.Utilities;

using dExcel.Utilities;
using NUnit.Framework;
using QLNet;

[TestFixture]
public class CommonUtilsTests
{
    [Test]
    [TestCase("Call", true, 1, null)]
    [TestCase("CALL", true, 1, null)]
    [TestCase("Put", true, -1, null)]
    [TestCase("PUT", true, -1, null)]
    [TestCase("Invalid", false, null, "#∂Excel Error: Invalid option type: 'Invalid'")]
    public void TryParseOptionTypeToSign_Test(
        string optionType, 
        bool expectedResult, 
        int? expectedSign, 
        string? expectedErrorMessage)
    {
        bool actualResult = 
            CommonUtils.TryParseOptionTypeToSign(optionType, out int? actualSign, out string? actualErrorMessage);
        
        Assert.AreEqual(expectedResult, actualResult);
        Assert.AreEqual(expectedSign, actualSign);
        Assert.AreEqual(expectedErrorMessage, actualErrorMessage);
    }

    [Test]
    [TestCase("Buy", true, 1, null)]
    [TestCase("BUY", true, 1, null)]
    [TestCase("Long", true, 1, null)]
    [TestCase("LONG", true, 1, null)]
    [TestCase("Sell", true, -1, null)]
    [TestCase("SELL", true, -1, null)]
    [TestCase("Short", true, -1, null)]
    [TestCase("SHORT", true, -1, null)]
    [TestCase("Invalid", false, null, "#∂Excel Error: Invalid direction: 'Invalid'")]
    public void TryParseDirectionToSign_Test(
        string direction, 
        bool expectedResult, 
        int? expectedSign, 
        string? expectedErrorMessage)
    {
        bool actualResult = 
            CommonUtils.TryParseDirectionToSign(direction, out int? actualSign, out string? actualErrorMessage);
        
        Assert.AreEqual(expectedResult, actualResult);
        Assert.AreEqual(expectedSign, actualSign);
        Assert.AreEqual(expectedErrorMessage, actualErrorMessage);
    }

    public static IEnumerable<TestCaseData> TryParseDayCountConvention_TestData()
    {
        yield return new TestCaseData("Act360", true, new Actual360(), null);     
        yield return new TestCaseData("ACT360", true, new Actual360(), null);     
        yield return new TestCaseData("Actual360", true, new Actual360(), null);     
        yield return new TestCaseData("ACTUAL360", true, new Actual360(), null);     
        yield return new TestCaseData("Act365", true, new Actual365Fixed(), null);
        yield return new TestCaseData("ACT365", true, new Actual365Fixed(), null);
        yield return new TestCaseData("ActAct", true, new ActualActual(), null);
        yield return new TestCaseData("ACTACT", true, new ActualActual(), null);
        yield return new TestCaseData("Business252", true, new Business252(), null);
        yield return new TestCaseData("BUSINESS252", true, new Business252(), null);
        yield return new TestCaseData("30360", true, new Thirty360(Thirty360.Thirty360Convention.BondBasis), null);
        yield return new TestCaseData("Thirty360", true, new Thirty360(Thirty360.Thirty360Convention.BondBasis), null);
        yield return new TestCaseData("Invalid", false, null, "#∂Excel Error: Invalid DayCountConvention: 'Invalid'");
    }
    
    [Test]
    [TestCaseSource(nameof(TryParseDayCountConvention_TestData))]
    public void TryParseDayCountConventionTest(
        string dayCountConventionToParse, 
        bool expectedResult, 
        DayCounter? expectedDayCountConvention, 
        string? expectedErrorMessage)
    {
        bool actualResult = 
            CommonUtils.TryParseDayCountConvention(
                dayCountConventionToParse: dayCountConventionToParse, 
                dayCountConvention: out DayCounter? actualDayCountConvention, 
                errorMessage: out string? actualErrorMessage); 
        
        Assert.AreEqual(expectedResult, actualResult);
        Assert.AreEqual(expectedDayCountConvention, actualDayCountConvention);
        Assert.AreEqual(expectedErrorMessage, actualErrorMessage);
    }
    
    public static IEnumerable<TestCaseData> TryParseInterpolation_TestData()
    {
        yield return new TestCaseData("BACKWARDFLAT", true, new BackwardFlat(), null);
        yield return new TestCaseData("CUBIC", true, new Cubic(), null);
        yield return new TestCaseData("FORWARDFLAT", true, new ForwardFlat(), null);
        yield return new TestCaseData("LINEAR", true, new Linear(), null);
        yield return new TestCaseData("LOGCUBIC", true, new LogCubic(), null);
        yield return new TestCaseData("EXPONENTIAL", true, new LogLinear(), null);
        yield return new TestCaseData("Invalid", false, null, "#∂Excel Error: Invalid interpolation method: 'Invalid'");
    }
    
    [Test]
    [TestCaseSource(nameof(TryParseInterpolation_TestData))]
    public void TryParseInterpolation_Test(
        string interpolationType, 
        bool expectedResult, 
        IInterpolationFactory? expectedInterpolation, 
        string? expectedErrorMessage)
    {
        bool actualResult = CommonUtils.TryParseInterpolation(
            interpolationMethodToParse: interpolationType,
            interpolation: out IInterpolationFactory? actualInterpolation, 
            errorMessage: out string? actualErrorMessage);  
        
        Assert.AreEqual(expectedResult, actualResult);
        Assert.AreEqual(expectedInterpolation?.GetType(), actualInterpolation?.GetType());
        Assert.AreEqual(expectedErrorMessage, actualErrorMessage);
    }
    
    public static IEnumerable<TestCaseData> TryParseCompoundingConvention_TestData()
    {
        yield return new TestCaseData("SIMPLE", true, (Compounding.Simple, Frequency.Once), null);
        yield return new TestCaseData("NACM", true, (Compounding.Compounded, Frequency.Monthly), null);
        yield return new TestCaseData("NACQ", true, (Compounding.Compounded, Frequency.Quarterly), null);
        yield return new TestCaseData("NACS", true, (Compounding.Compounded, Frequency.Semiannual), null);
        yield return new TestCaseData("NACA", true, (Compounding.Compounded, Frequency.Annual), null);
        yield return new TestCaseData("NACC", true, (Compounding.Continuous, Frequency.NoFrequency), null);
        yield return new TestCaseData("Invalid", false, null, "#∂Excel Error: Invalid compounding convention: 'Invalid'");
    }
    
    [Test]
    [TestCaseSource(nameof(TryParseCompoundingConvention_TestData))]
    public void TryParseCompoundingConvention_Test(
        string compoundingConvention, 
        bool expectedResult, 
        (Compounding, Frequency)? expectedCompoundingConvention, 
        string? expectedErrorMessage)
    {
        bool actualResult = CommonUtils.TryParseCompoundingConvention(
            compoundingConventionToParse: compoundingConvention,
            compoundingConvention: out (Compounding, Frequency)? actualCompoundingConvention, 
            errorMessage: out string? actualErrorMessage);  
        
        Assert.AreEqual(expectedResult, actualResult);
        Assert.AreEqual(expectedCompoundingConvention?.GetType(), actualCompoundingConvention?.GetType());
        Assert.AreEqual(expectedErrorMessage, actualErrorMessage);
    }
}
