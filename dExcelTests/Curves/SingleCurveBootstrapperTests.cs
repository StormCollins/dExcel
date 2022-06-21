namespace dExcelTests;

using dExcel;
using dExcel.Curves;
using NUnit.Framework;
using QLNet;

[TestFixture]
public class SingleCurveBootstrapperTests
{
    [Test]
    public void BootstrapFlatCurveDepositsTest()
    {
        DateTime baseDate = new DateTime(2022, 06, 01);
        object[,] instruments = 
        {
            {"Deposits", "", "", ""},
            {"Tenors", "RateIndex", "Rates", "Include"},
            {"1m", "JIBAR", 0.1, "TRUE"},
            {"3m", "JIBAR", 0.1, "TRUE"},
            {"6m", "JIBAR", 0.1, "TRUE"},
        };
        
        var dayCounter = new Actual365Fixed();
        var handle = SingleCurveBootstrapper.Bootstrap("BootstrappedSingleCurve", baseDate, instruments);
        var curve = (YieldTermStructure)((Dictionary<string, object>)DataObjectController.GetDataObject(handle))["Curve"];
        const double tolerance = 0.001; 
        
        Assert.AreEqual(1.0, curve.discount(baseDate));
        Assert.AreEqual(
            expected: Math.Exp(-0.1 * dayCounter.yearFraction(baseDate, baseDate.AddMonths(1))),
            actual: curve.discount(baseDate.AddMonths(1)),
            delta: tolerance);
        Assert.AreEqual(
            expected: Math.Exp(-0.1 * dayCounter.yearFraction(baseDate, baseDate.AddMonths(2))),
            actual: curve.discount(baseDate.AddMonths(2)),
            delta: tolerance);
        Assert.AreEqual(
            expected: Math.Exp(-0.1 * dayCounter.yearFraction(baseDate, baseDate.AddMonths(3))),
            actual: curve.discount(baseDate.AddMonths(3)),
            delta: tolerance);
        Assert.AreEqual(
            expected: Math.Exp(-0.1 * dayCounter.yearFraction(baseDate, baseDate.AddMonths(6))),
            actual: curve.discount(baseDate.AddMonths(6)),
            delta: tolerance);
    }
    
    [Test]
    public void BootstrapFlatCurveDepositsAndFrasTest()
    {
        DateTime baseDate = new DateTime(2022, 06, 01);
        object[,] depositInstruments = 
        {
            {"Deposits", "", "", ""},
            {"Tenors", "RateIndex", "Rates", "Include"},
            {"1m", "JIBAR", 0.1, "TRUE"},
            {"3m", "JIBAR", 0.1, "TRUE"},
            {"6m", "JIBAR", 0.1, "TRUE"},
        };
       
        object[,] fraInstruments = 
        {
            {"FRAs", "", "", ""},
            {"Tenors", "RateIndex", "Rates", "Include"},
            {"6m", "JIBAR", 0.1, "TRUE"},
            {"9m", "JIBAR", 0.1, "TRUE"},
        };

        object[] instruments = {depositInstruments, fraInstruments};
        var dayCounter = new Actual365Fixed();
        var handle = SingleCurveBootstrapper.Bootstrap("BootstrappedSingleCurve", baseDate, instruments);
        var curve = (YieldTermStructure)((Dictionary<string, object>)DataObjectController.GetDataObject(handle))["Curve"];
        const double tolerance = 0.01; 
        
        Assert.AreEqual(1.0, curve.discount(baseDate));
        Assert.AreEqual(
            expected: Math.Exp(-0.1 * dayCounter.yearFraction(baseDate, baseDate.AddMonths(1))),
            actual: curve.discount(baseDate.AddMonths(1)),
            delta: tolerance);
        Assert.AreEqual(
            expected: Math.Exp(-0.1 * dayCounter.yearFraction(baseDate, baseDate.AddMonths(2))),
            actual: curve.discount(baseDate.AddMonths(2)),
            delta: tolerance);
        Assert.AreEqual(
            expected: Math.Exp(-0.1 * dayCounter.yearFraction(baseDate, baseDate.AddMonths(3))),
            actual: curve.discount(baseDate.AddMonths(3)),
            delta: tolerance);
        Assert.AreEqual(
            expected: Math.Exp(-0.1 * dayCounter.yearFraction(baseDate, baseDate.AddMonths(6))),
            actual: curve.discount(baseDate.AddMonths(6)),
            delta: tolerance);
        Assert.AreEqual(
            expected: Math.Exp(-0.1 * dayCounter.yearFraction(baseDate, baseDate.AddMonths(6))) *
                      Math.Exp(-0.1 * dayCounter.yearFraction(baseDate.AddMonths(6), baseDate.AddMonths(9))),
            actual: curve.discount(baseDate.AddMonths(9)),
            delta: tolerance);
    }
}
