namespace dExcelTests.Utilities;

using dExcel.Utilities;
using NUnit.Framework;
using QLNet;

[TestFixture]
public class CommonUtilsTests
{
    // Since we have to test the "out" parameters as well, there doesn't seem to be an elegant, easily readable way to
    // do this, other than just testing each case in turn.
    [Test]
    [TestCase("Act360")]
    [TestCase("ACT360")]
    [TestCase("Actual360")]
    [TestCase("ACTUAL360")]
    public void TryParseDayCountConventionForActual360Test(string dayCountConventionToParse)
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
    public void TryParseDayCountConventionForActual365Test(string dayCountConventionToParse)
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
    public void TryParseDayCountConventionForActualActualTest(string dayCountConventionToParse)
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
    public void TryParseDayCountConventionForBusiness252Test(string dayCountConventionToParse)
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
    public void TryParseDayCountConvention30360Test(string dayCountConventionToParse)
    {
        bool actual = CommonUtils.TryParseDayCountConvention(dayCountConventionToParse, out DayCounter? actualDayCountConvention, out string? errorMessage); 
        Assert.IsTrue(actual);
        Assert.AreEqual(new Thirty360(Thirty360.Thirty360Convention.BondBasis), actualDayCountConvention);
        Assert.AreEqual(null, errorMessage);
    }
    
    // Since we have to test the "out" parameters as well, there doesn't seem to be an elegant, easily readable way to
    // do this, other than just testing each case in turn.
    [Test]
    public void TryParseDayCountConventionErrorMessageTest()
    {
        bool actual = CommonUtils.TryParseDayCountConvention("Invalid", out DayCounter? actualDayCountConvention, out string? errorMessage); 
        Assert.IsFalse(actual);
        Assert.AreEqual(null, actualDayCountConvention);
        Assert.AreEqual(CommonUtils.DExcelErrorMessage("Invalid DayCountConvention: 'Invalid'"), errorMessage);
    }
}
