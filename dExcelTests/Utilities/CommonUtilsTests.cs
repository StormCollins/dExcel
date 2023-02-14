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
    public void TryParseOptionTypeToSign_Test(string optionType, bool expectedResult, int? expectedSign, string? expectedErrorMessage)
    {
        bool actualResult = 
            CommonUtils.TryParseOptionTypeToSign(optionType, out int? actualSign, out string? actualErrorMessage);
        Assert.AreEqual(expectedResult, actualResult);
        Assert.AreEqual(expectedSign, actualSign);
        Assert.AreEqual(expectedErrorMessage, actualErrorMessage);
    }
    
    
    // Since we have to test the "out" parameters as well, there doesn't seem to be an elegant, easily readable way to
    // do this, other than just testing each case in turn.
    [Test]
    [TestCase("Act360")]
    [TestCase("ACT360")]
    [TestCase("Actual360")]
    [TestCase("ACTUAL360")]
    public void TryParseDayCountConvention_Actual360Test(string dayCountConventionToParse)
    {
        bool actual = CommonUtils.TryParseDayCountConvention(dayCountConventionToParse, out DayCounter? actualDayCountConvention, out string? errorMessage); 
        Assert.IsTrue(actual);
        Assert.AreEqual(new Actual360(), actualDayCountConvention);
        Assert.AreEqual(null, errorMessage);
    }
    
    // Since we have to test the "out" parameters as well, there doesn't seem to be an elegant, easily readable way to
    // do this, other than just testing each case in turn.
    [Test]
    [TestCase("Act365")]
    [TestCase("ACT365")]
    [TestCase("Actual365")]
    [TestCase("ACTUAL365")]
    public void TryParseDayCountConvention_Actual365Test(string dayCountConventionToParse)
    {
        bool actual = CommonUtils.TryParseDayCountConvention(dayCountConventionToParse, out DayCounter? actualDayCountConvention, out string? errorMessage); 
        Assert.IsTrue(actual);
        Assert.AreEqual(new Actual365Fixed(), actualDayCountConvention);
        Assert.AreEqual(null, errorMessage);
    }
    
    // Since we have to test the "out" parameters as well, there doesn't seem to be an elegant, easily readable way to
    // do this, other than just testing each case in turn.
    [Test]
    [TestCase("ActAct")]
    [TestCase("ActAct")]
    [TestCase("ActualActual")]
    [TestCase("ACTUALACTUAL")]
    public void TryParseDayCountConvention_ActualActualTest(string dayCountConventionToParse)
    {
        bool actual = CommonUtils.TryParseDayCountConvention(dayCountConventionToParse, out DayCounter? actualDayCountConvention, out string? errorMessage); 
        Assert.IsTrue(actual);
        Assert.AreEqual(new ActualActual(), actualDayCountConvention);
        Assert.AreEqual(null, errorMessage);
    }
    
    // Since we have to test the "out" parameters as well, there doesn't seem to be an elegant, easily readable way to
    // do this, other than just testing each case in turn.
    [Test]
    [TestCase("Business252")]
    [TestCase("BUSINESS252")]
    public void TryParseDayCountConvention_Business252Test(string dayCountConventionToParse)
    {
        bool actual = CommonUtils.TryParseDayCountConvention(dayCountConventionToParse, out DayCounter? actualDayCountConvention, out string? errorMessage); 
        Assert.IsTrue(actual);
        Assert.AreEqual(new Business252(), actualDayCountConvention);
        Assert.AreEqual(null, errorMessage);
    }
    
    // Since we have to test the "out" parameters as well, there doesn't seem to be an elegant, easily readable way to
    // do this, other than just testing each case in turn.
    [Test]
    [TestCase("30360")]
    [TestCase("Thirty360")]
    [TestCase("THIRTY360")]
    public void TryParseDayCountConvention_30360Test(string dayCountConventionToParse)
    {
        bool actual = CommonUtils.TryParseDayCountConvention(dayCountConventionToParse, out DayCounter? actualDayCountConvention, out string? errorMessage); 
        Assert.IsTrue(actual);
        Assert.AreEqual(new Thirty360(Thirty360.Thirty360Convention.BondBasis), actualDayCountConvention);
        Assert.AreEqual(null, errorMessage);
    }
    
    // Since we have to test the "out" parameters as well, there doesn't seem to be an elegant, easily readable way to
    // do this, other than just testing each case in turn.
    [Test]
    public void TryParseDayCountConvention_ErrorMessageTest()
    {
        bool actual = CommonUtils.TryParseDayCountConvention("Invalid", out DayCounter? actualDayCountConvention, out string? errorMessage); 
        Assert.IsFalse(actual);
        Assert.AreEqual(null, actualDayCountConvention);
        Assert.AreEqual(CommonUtils.DExcelErrorMessage("Invalid DayCountConvention: 'Invalid'"), errorMessage);
    }
}
