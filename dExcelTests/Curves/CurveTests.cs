using dExcel.InterestRates;

namespace dExcelTests.Curves;

using dExcel.Curves;
using NUnit.Framework;

[TestFixture]
public class CurveTests
{
    private static readonly object[,] CurveParameters =
    {
        { "Parameter", "Value" },
        { "DayCountConvention", "Actual365" },
        { "Interpolation", "Exponential" },
    };

    private static readonly object[,] DatesRange =
        {
            { new DateTime(2022, 01, 01).ToOADate() },
            { new DateTime(2022, 02, 01).ToOADate() },
            { new DateTime(2022, 03, 01).ToOADate() },
            { new DateTime(2022, 04, 01).ToOADate() },
            { new DateTime(2022, 05, 01).ToOADate() },
            { new DateTime(2025, 12, 31).ToOADate() },
        };

    private static readonly object[,] DiscountFactorRange =
    {
        { 1.00 },
        { 0.99 },
        { 0.98 },
        { 0.97 },
        { 0.96 },
        { 0.62 },
    };

    private static readonly string Handle = CurveUtils.Create("ZarSwapCurve", CurveParameters, DatesRange, DiscountFactorRange);
        
    [TestCase]
    public void TestZarGetZeroRates()
    {
        object[,] dates = 
        {
            { new DateTime(2022, 02, 01).ToOADate() },
            { new DateTime(2022, 03, 01).ToOADate() },
            { new DateTime(2022, 04, 01).ToOADate() },
            { new DateTime(2022, 05, 01).ToOADate() },
            { new DateTime(2025, 12, 31).ToOADate() },
        };

        object[,] zeroRates = (object[,])CurveUtils.GetZeroRates(Handle, dates);
        
        // The derivation of the "actual" figures can be found in the sheet "Curves" of the workbook
        // "dexcel-testing.xlsm".
        Assert.AreEqual(0.118335, (double)zeroRates[0, 0], 0.00001);
        Assert.AreEqual(0.124983, (double)zeroRates[1, 0], 0.00001);
        Assert.AreEqual(0.123529, (double)zeroRates[2, 0], 0.00001);
        Assert.AreEqual(0.124167, (double)zeroRates[3, 0], 0.00001);
        Assert.AreEqual(0.119509, (double)zeroRates[4, 0], 0.00001);
    }
    
    [TestCase]
    public void TestZarGetDiscountFactors()
    {
        object[] interpolationDates = new object[] { new DateTime(2022, 01, 15).ToOADate() };
        object[,] dfs = CurveUtils.GetDiscountFactors(Handle, interpolationDates);

        double expectedDf =
            Math.Exp((Math.Log(0.99) - Math.Log(1.00))
            / (((double)DatesRange[1, 0] - (double)DatesRange[0, 0]) / 365)
            * (((double)interpolationDates[0] - (double)DatesRange[0, 0]) / 365)
            + Math.Log(1.00));

        Assert.AreEqual(expectedDf, dfs[0, 0]);
    }
}
