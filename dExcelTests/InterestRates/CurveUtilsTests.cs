using dExcel.CommonEnums;
using dExcel.InterestRates;
using dExcel.Utilities;
using NUnit.Framework;

namespace dExcelTests.InterestRates;

using dExcel.Dates;
using QuantLib;

[TestFixture]
public class CurveUtilsTests
{
    private static readonly object[,] CurveParameters =
    {
        { "Parameter", "Value" },
        { "Calendars", "ZAR" },
        { "DayCountConvention", "Actual365" },
        { "Interpolation", CurveInterpolationMethods.Exponential_On_DiscountFactors.ToString() },
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

    private static readonly object[,] IncompatiblySizedDiscountFactorRange =
    {
        { 1.00 },
        { 0.99 },
        { 0.98 },
        { 0.97 },
        { 0.96 },
    };

    private static readonly string Handle = CurveUtils.CreateFromDiscountFactors("ZarSwapCurve", CurveParameters, DatesRange, DiscountFactorRange);
    
     
    [TestCase]
    public void ZarGetZeroRatesTest()
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
    
    [Test]
    public void ZarGetDiscountFactorsTest()
    {
        string handle1 = CurveUtils.CreateFromDiscountFactors("ZarSwapCurve", CurveParameters, DatesRange, DiscountFactorRange); 
        object[] interpolationDates = { new DateTime(2022, 01, 15).ToOADate() };
        object[,] dfs = (object[,])CurveUtils.GetDiscountFactors(handle1, interpolationDates);

        double expectedDf =
            Math.Exp((Math.Log(0.99) - Math.Log(1.00))
            / (((double)DatesRange[1, 0] - (double)DatesRange[0, 0]) / 365)
            * (((double)interpolationDates[0] - (double)DatesRange[0, 0]) / 365)
            + Math.Log(1.00));

        Assert.AreEqual(expectedDf, dfs[0, 0]);
    }

    [Test]
    public void TestIncompatiblySizedDateAndDiscountFactorsRange()
    {
        string actual = 
            CurveUtils.CreateFromDiscountFactors("ErrorCurve", CurveParameters, DatesRange, IncompatiblySizedDiscountFactorRange);
        
        string expected = 
            CommonUtils.DExcelErrorMessage(
                $"Dates and discount factors have incompatible sizes: " +
                $"({DatesRange.GetLength(0)} ≠ {IncompatiblySizedDiscountFactorRange.GetLength(0)}).");
        
        Assert.AreEqual(expected, actual);
    }
    
    
    [Test]
    public void GetRateIndicesTest()
    {
        Array rateIndices = Enum.GetValues(typeof(RateIndices));
        object[,] expectedOutput = new object[rateIndices.Length + 1, 1];
        expectedOutput[0, 0] = "Rate Indices";
        int i = 1;
        foreach (RateIndices index in rateIndices)
        {
            expectedOutput[i++, 0] = index.ToString();
        }
        object[,] actualOutput = (object[,])CurveUtils.GetRateIndices();
        
        Assert.AreEqual(expectedOutput, actualOutput);
    }
    
    [Test]
    public void GetInterpolationMethodsForDiscountFactorsTest()
    {
        List<string> interpolationMethodsForDiscountFactors = 
            Enum.GetNames(typeof(CurveInterpolationMethods))
                .Where(x => x.ToUpper().Contains("DISCOUNTFACTORS"))
                .ToList();
        
        object[,] expectedOutput = new object[interpolationMethodsForDiscountFactors.Count + 1, 1];
        expectedOutput[0, 0] = "Interpolation Methods for Discount Factors";
        int i = 1;
        foreach (string curveInterpolationMethod in interpolationMethodsForDiscountFactors)
        {
            expectedOutput[i++, 0] = curveInterpolationMethod;
        }
        
        object[,] actualOutput = (object[,])CurveUtils.GetInterpolationMethodsForDiscountFactors();
        Assert.AreEqual(expectedOutput, actualOutput);
    }
    
    [Test]
    public void GetInterpolationMethodsForZeroRatesTest()
    {
        List<string> interpolationMethods = 
            Enum.GetNames(typeof(CurveInterpolationMethods))
                .Where(x => x.ToUpper().Contains("ZERORATES"))
                .ToList();
        
        object[,] expectedOutput = new object[interpolationMethods.Count + 1, 1];
        expectedOutput[0, 0] = "Interpolation Methods for Zero Rates";
        int i = 1;
        foreach (string interpolationMethod in interpolationMethods)
        {
            expectedOutput[i++, 0] = interpolationMethod;
        }
        
        object[,] actualOutput = (object[,])CurveUtils.GetInterpolationMethodsForZeroRates();
        Assert.AreEqual(expectedOutput, actualOutput);
    }

    [Test]
    public void CreateFromZeroRatesTest()
    {
        DateTime baseDate = DateTime.FromOADate((double) DatesRange[0, 0]);
        object[,] yearFractions = new object[DatesRange.GetLength(0), 1];
        for (int i = 0; i < yearFractions.GetLength(0); i++)
        {
            yearFractions[i, 0] = DateUtils.Act365(baseDate, DateTime.FromOADate((double) DatesRange[i, 0]));
        }  
        
        object[,] zeroRates = new object[DiscountFactorRange.GetLength(0), 1];
        for (int i = 0; i < zeroRates.GetLength(0); i++)
        {
            zeroRates[i, 0] = -1 * Math.Log((double) DiscountFactorRange[i, 0]) / (double) yearFractions[i, 0];
        }

        object[,] curveParameters =
        {
            {"Parameter", "Value"},
            {"BaseDate", (double) DatesRange[0, 0]},
            {"Calendars", "ZAR"},
            {"DayCountConvention", "Actual365"},
            {"CompoundingConvention", "NACC"},
            {"Interpolation", CurveInterpolationMethods.Linear_On_ZeroRates.ToString()},
        };
        
        string handle = 
            CurveUtils.CreateFromZeroRates(nameof(CreateFromZeroRatesTest), curveParameters, DatesRange, zeroRates);

        for (int i = 0; i < DiscountFactorRange.GetLength(0); i++)
        {
            Assert.AreEqual(
                expected: (double)DiscountFactorRange[i, 0], 
                actual: (double)((object[,])CurveUtils.GetDiscountFactors(handle, new[] {DatesRange[i, 0]}))[0, 0],
                delta: 1e-10); 
        }  
    }
}
