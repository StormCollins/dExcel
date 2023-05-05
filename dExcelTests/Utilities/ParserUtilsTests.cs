using dExcel.Utilities;
using NUnit.Framework;
using QL = QuantLib;

namespace dExcelTests.Utilities;

[TestFixture]
public class ParserUtilsTests
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
            ParserUtils.TryParseOptionTypeToSign(optionType, out int? actualSign, out string? actualErrorMessage);
        
        Assert.AreEqual(expectedResult, actualResult);
        Assert.AreEqual(expectedSign, actualSign);
        Assert.AreEqual(expectedErrorMessage, actualErrorMessage);
    }

    [Test]
    [TestCase("Call", true, QL.Option.Type.Call, null)]
    [TestCase("CALL", true, QL.Option.Type.Call, null)]
    [TestCase("Put", true, QL.Option.Type.Put, null)]
    [TestCase("PUT", true, QL.Option.Type.Put, null)]
    [TestCase("Invalid", false, null, "#∂Excel Error: Invalid option type: 'Invalid'")]
    public void TryParseOptionTypeToQuantLibType_Test(
        string optionType,
        bool expectedResult,
        QL.Option.Type? expected,
        string? expectedErrorMessage)
    {
        bool actualResult =
            ParserUtils.TryParseQuantLibOptionType(optionType, out QL.Option.Type? quantLibOptionType, out string? actualErrorMessage);

        Assert.AreEqual(expectedResult, actualResult);
        Assert.AreEqual(expected, quantLibOptionType);
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
            ParserUtils.TryParseDirectionToSign(direction, out int? actualSign, out string? actualErrorMessage);
        
        Assert.AreEqual(expectedResult, actualResult);
        Assert.AreEqual(expectedSign, actualSign);
        Assert.AreEqual(expectedErrorMessage, actualErrorMessage);
    }

    public static IEnumerable<TestCaseData> TryParseDayCountConvention_TestData()
    {
        yield return new TestCaseData("Act360", true, new QL.Actual360(), null);     
        yield return new TestCaseData("ACT360", true, new QL.Actual360(), null);     
        yield return new TestCaseData("Actual360", true, new QL.Actual360(), null);     
        yield return new TestCaseData("ACTUAL360", true, new QL.Actual360(), null);     
        yield return new TestCaseData("Act365", true, new QL.Actual365Fixed(), null);
        yield return new TestCaseData("ACT365", true, new QL.Actual365Fixed(), null);
        yield return new TestCaseData("ActAct", true, new QL.ActualActual(QL.ActualActual.Convention.ISDA), null);
        yield return new TestCaseData("ACTACT", true, new QL.ActualActual(QL.ActualActual.Convention.ISDA), null);
        yield return new TestCaseData("Business252", true, new QL.Business252(), null);
        yield return new TestCaseData("BUSINESS252", true, new QL.Business252(), null);
        yield return new TestCaseData("30360", true, new QL.Thirty360(QL.Thirty360.Convention.ISDA), null);
        yield return new TestCaseData("Thirty360", true, new QL.Thirty360(QL.Thirty360.Convention.ISDA), null);
        yield return new TestCaseData("Invalid", false, null, "#∂Excel Error: Invalid DayCountConvention: 'Invalid'");
    }
    
    [Test]
    [TestCaseSource(nameof(TryParseDayCountConvention_TestData))]
    public void TryParseDayCountConventionTest(
        string dayCountConventionToParse, 
        bool expectedResult, 
        QL.DayCounter? expectedDayCountConvention, 
        string? expectedErrorMessage)
    {
        bool actualResult = 
            ParserUtils.TryParseQuantLibDayCountConvention(
                dayCountConventionToParse: dayCountConventionToParse, 
                dayCountConvention: out QL.DayCounter? actualDayCountConvention, 
                errorMessage: out string? actualErrorMessage); 
        
        Assert.AreEqual(expectedResult, actualResult);
        if (expectedDayCountConvention != null)
            Assert.AreEqual(expectedDayCountConvention.name(), actualDayCountConvention.name());
        Assert.AreEqual(expectedErrorMessage, actualErrorMessage);
    }
    
    public static IEnumerable<TestCaseData> TryParseCompoundingConvention_TestData()
    {
        yield return new TestCaseData("SIMPLE", true, (QL.Compounding.Simple, QL.Frequency.Once), null);
        yield return new TestCaseData("NACM", true, (QL.Compounding.Compounded, QL.Frequency.Monthly), null);
        yield return new TestCaseData("NACQ", true, (QL.Compounding.Compounded, QL.Frequency.Quarterly), null);
        yield return new TestCaseData("NACS", true, (QL.Compounding.Compounded, QL.Frequency.Semiannual), null);
        yield return new TestCaseData("NACA", true, (QL.Compounding.Compounded, QL.Frequency.Annual), null);
        yield return new TestCaseData("NACC", true, (QL.Compounding.Continuous, QL.Frequency.NoFrequency), null);
        yield return new TestCaseData("Invalid", false, null, "#∂Excel Error: Invalid compounding convention: 'Invalid'");
    }
    
    [Test]
    [TestCaseSource(nameof(TryParseCompoundingConvention_TestData))]
    public void TryParseCompoundingConvention_Test(
        string compoundingConvention, 
        bool expectedResult, 
        (QL.Compounding, QL.Frequency)? expectedCompoundingConvention, 
        string? expectedErrorMessage)
    {
        bool actualResult = ParserUtils.TryParseQuantLibCompoundingConvention(
            compoundingConventionToParse: compoundingConvention,
            compoundingConvention: out (QL.Compounding, QL.Frequency)? actualCompoundingConvention, 
            errorMessage: out string? actualErrorMessage);  
        
        Assert.AreEqual(expectedResult, actualResult);
        Assert.AreEqual(expectedCompoundingConvention?.GetType(), actualCompoundingConvention?.GetType());
        Assert.AreEqual(expectedErrorMessage, actualErrorMessage);
    }
}
