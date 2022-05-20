namespace dExcelTests;

using dExcel.Curves;
using NUnit.Framework;

[TestFixture]
public class CurveTests
{
    [TestCase]
    public void TestZarDiscountFactors()
    {
        object[,] datesRange =
            {
                { new DateTime(2022, 01, 01).ToOADate() },
                { new DateTime(2022, 02, 01).ToOADate() },
                { new DateTime(2022, 03, 01).ToOADate() },
                { new DateTime(2022, 04, 01).ToOADate() },
                { new DateTime(2022, 05, 01).ToOADate() }
            };

        object[,] discountFactorRange =
        {
            { 1.00 },
            { 0.99 },
            { 0.98 },
            { 0.97 },
            { 0.96 },
        };

        string handle = Curve.Create("ZarSwapCurve", datesRange, discountFactorRange);

        var interpolationDates = new object[] { new DateTime(2022, 01, 15).ToOADate() };
        object[,] dfs = Curve.GetDiscountFactors(handle, interpolationDates);

        double expectedDf =
            Math.Exp((Math.Log(0.99) - Math.Log(1.00))
            / (((double)datesRange[1, 0] - (double)datesRange[0, 0]) / 365)
            * (((double)interpolationDates[0] - (double)datesRange[0, 0]) / 365)
            + Math.Log(1.00));

        Assert.AreEqual(expectedDf, dfs[0, 0]);
    }
}
