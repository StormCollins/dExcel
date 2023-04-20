﻿using dExcel.Utilities;
using NUnit.Framework;
using QL = QuantLib;

namespace dExcelTests.Utilities;

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
            CommonUtils.TryParseDayCountConvention(
                dayCountConventionToParse: dayCountConventionToParse, 
                dayCountConvention: out QL.DayCounter? actualDayCountConvention, 
                errorMessage: out string? actualErrorMessage); 
        
        Assert.AreEqual(expectedResult, actualResult);
        Assert.AreEqual(expectedDayCountConvention, actualDayCountConvention);
        Assert.AreEqual(expectedErrorMessage, actualErrorMessage);
    }
    
    // public static IEnumerable<TestCaseData> TryParseInterpolation_TestData()
    // {
    //     yield return new TestCaseData("BACKWARDFLAT", true, new BackwardFlat(), null);
    //     yield return new TestCaseData("CUBIC", true, new Cubic(), null);
    //     yield return new TestCaseData("FORWARDFLAT", true, new ForwardFlat(), null);
    //     yield return new TestCaseData("LINEAR", true, new Linear(), null);
    //     yield return new TestCaseData("LOGCUBIC", true, new LogCubic(), null);
    //     yield return new TestCaseData("EXPONENTIAL", true, new LogLinear(), null);
    //     yield return new TestCaseData("Invalid", false, null, "#∂Excel Error: Invalid interpolation method: 'Invalid'");
    // }
    //
    // [Test]
    // [TestCaseSource(nameof(TryParseInterpolation_TestData))]
    // public void TryParseInterpolation_Test(
    //     string interpolationType, 
    //     bool expectedResult, 
    //     IInterpolationFactory? expectedInterpolation, 
    //     string? expectedErrorMessage)
    // {
    //     bool actualResult = CommonUtils.TryParseInterpolation(
    //         interpolationMethodToParse: interpolationType,
    //         interpolation: out IInterpolationFactory? actualInterpolation, 
    //         errorMessage: out string? actualErrorMessage);  
    //     
    //     Assert.AreEqual(expectedResult, actualResult);
    //     Assert.AreEqual(expectedInterpolation?.GetType(), actualInterpolation?.GetType());
    //     Assert.AreEqual(expectedErrorMessage, actualErrorMessage);
    // }
    
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
        bool actualResult = CommonUtils.TryParseCompoundingConvention(
            compoundingConventionToParse: compoundingConvention,
            compoundingConvention: out (QL.Compounding, QL.Frequency)? actualCompoundingConvention, 
            errorMessage: out string? actualErrorMessage);  
        
        Assert.AreEqual(expectedResult, actualResult);
        Assert.AreEqual(expectedCompoundingConvention?.GetType(), actualCompoundingConvention?.GetType());
        Assert.AreEqual(expectedErrorMessage, actualErrorMessage);
    }
}
