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
        DateTime baseDate = new(2022, 06, 01);

        object[,] curveParameters =
        {
            {"Curve Parameters", ""},
            {"Parameter", "Value"},
            {"BaseDate", baseDate.ToOADate()},
            {"RateIndex", "JIBAR"},
        };

        object[,] instruments = 
        {
            {"Deposits", "", "", ""},
            {"Tenors", "RateIndex", "Rates", "Include"},
            {"1m", "JIBAR", 0.1, "TRUE"},
            {"3m", "JIBAR", 0.1, "TRUE"},
            {"6m", "JIBAR", 0.1, "TRUE"},
        };
        
        var dayCounter = new Actual365Fixed();
        var handle = SingleCurveBootstrapper.Bootstrap("BootstrappedSingleCurve", curveParameters, instruments);
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
        DateTime baseDate = new(2022, 06, 01);
        
        object[,] curveParameters =
        {
            {"Curve Parameters", ""},
            {"Parameter", "Value"},
            {"BaseDate", baseDate.ToOADate()},
            {"RateIndex", "JIBAR"},
        };

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
        var handle = SingleCurveBootstrapper.Bootstrap("BootstrappedSingleCurve", curveParameters, null, instruments);
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
    
    [Test]
    public void BootstrapFlatCurveDepositsFrasAndSwapsTest()
    {
        var baseDate = new DateTime(2022, 06, 01);

        object[,] curveParameters =
        {
            {"Curve Parameters", ""},
            {"Parameter", "Value"},
            {"BaseDate", baseDate.ToOADate()},
            {"RateIndex", "JIBAR"},
        };

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

        object[,] swapInstruments =
        {
            {"Interest Rate Swaps", "", "", ""},
            {"Tenors", "RateIndex", "Rates", "Include"},
            {"1y", "JIBAR", 0.1, "TRUE"},
            {"2y", "JIBAR", 0.1, "TRUE"},
        };

        object[] instruments = {depositInstruments, fraInstruments, swapInstruments};
        var dayCounter = new Actual365Fixed();
        var handle = SingleCurveBootstrapper.Bootstrap("BootstrappedSingleCurve", curveParameters, null, instruments);
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
