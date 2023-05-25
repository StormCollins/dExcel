using System.Diagnostics;
using dExcel.CommonEnums;
using dExcel.Dates;
using dExcel.ExcelUtils;
using dExcel.InterestRates;
using NUnit.Framework;
using QL = QuantLib;
using DateTime = System.DateTime;

namespace dExcelTests.InterestRates;

[TestFixture]
public class CurveBootstrapperTest
{
    [Test]
    public void Bootstrap_MissingBaseDate_Test()
    {
        object[,] curveParameters =
        {
            {"CurveUtils Parameters", ""},
            {"Parameter", "Value"},
            {"RateIndexName", RateIndices.JIBAR.ToString()},
            {"RateIndexTenor", "3m"},
            {"Interpolation", CurveInterpolationMethods.Exponential_On_DiscountFactors.ToString()},
        };

        object[,] instrumentGroups = 
        {
            {"Deposits", "", "", ""},
            {"Tenors", "RateIndex", "Rates", "Include"},
            {"1m", RateIndices.JIBAR.ToString(), 0.1, "TRUE"},
            {"3m", RateIndices.JIBAR.ToString(), 0.1, "TRUE"},
            {"6m", RateIndices.JIBAR.ToString(), 0.1, "TRUE"},
        };
        
        string handle = 
            CurveBootstrapper.Bootstrap(
                handle: "BootstrappedSingleCurve", 
                curveParameters: curveParameters,
                customRateIndex: null,
                instrumentGroups: instrumentGroups);
        
            const string expected = "#∂Excel Error: Missing curve parameter: 'BASEDATE'.";
            Assert.AreEqual(expected, handle);
    }
        
    [Test]
    public void Bootstrap_FlatCurve_Deposits_Test()
    {
        DateTime baseDate = new(2022, 06, 01);

        object[,] curveParameters =
        {
            {"CurveUtils Parameters", ""},
            {"Parameter", "Value"},
            {"BaseDate", baseDate.ToOADate()},
            {"RateIndexName", "JIBAR"},
            {"RateIndexTenor", "3m"},
            {"Interpolation", CurveInterpolationMethods.Exponential_On_DiscountFactors.ToString()},
        };

        Dictionary<string, double> depositRates = 
            new()
            {
                ["1m"] = 0.1,
                ["3m"] = 0.1,
                ["6m"] = 0.1,
            };
        
        object[,] instruments = 
        {
            {"Deposits", "", "", ""},
            {"Tenors", "RateIndex", "Rates", "Include"},
            {"1m", RateIndices.JIBAR.ToString(), depositRates["1m"], "TRUE"},
            {"3m", RateIndices.JIBAR.ToString(), depositRates["3m"], "TRUE"},
            {"6m", RateIndices.JIBAR.ToString(), depositRates["6m"], "TRUE"},
        };
        
        QL.DayCounter dayCounter = new QL.Actual365Fixed();
        string handle = 
            CurveBootstrapper.Bootstrap(
                handle: "BootstrappedSingleCurve", 
                curveParameters, 
                customRateIndex: null,
                instruments);

        QL.YieldTermStructure? curve = CurveUtils.GetCurveObject(handle);
        const double tolerance = 0.01;
        Debug.Assert(curve != null, nameof(curve) + " != null");
        Assert.AreEqual(1.0, curve.discount(baseDate.ToQuantLibDate()));
                    
        DateTime date1M = (DateTime) DateUtils.AddTenorToDate(baseDate, "1m", "ZAR", "ModFol");
        double discountFactor1M = 1 / (1 + depositRates["1m"] * dayCounter.yearFraction(baseDate.ToQuantLibDate(), date1M.ToQuantLibDate()));
        Assert.AreEqual(discountFactor1M, curve.discount(date1M.ToQuantLibDate()), tolerance);
        
        DateTime date3M = (DateTime) DateUtils.AddTenorToDate(baseDate, "3m", "ZAR", "ModFol");
        double discountFactor3M = 1 / (1 + depositRates["3m"] * dayCounter.yearFraction(baseDate.ToQuantLibDate(), date3M.ToQuantLibDate()));
        Assert.AreEqual(discountFactor3M, curve.discount(date3M.ToQuantLibDate()), tolerance);
        
        DateTime date6M = (DateTime) DateUtils.AddTenorToDate(baseDate, "6m", "ZAR", "ModFol");
        double discountFactor6M = 1 / (1 + depositRates["6m"] * dayCounter.yearFraction(baseDate.ToQuantLibDate(), date6M.ToQuantLibDate()));
        Assert.AreEqual(discountFactor6M, curve.discount(date6M.ToQuantLibDate()), tolerance); 
    }
    
    [Test]
    public void Bootstrap_FlatCurve_DepositsAndFras_Test()
    {
        DateTime baseDate = new(2022, 06, 01);
        
        object[,] curveParameters =
        {
            {"CurveUtils Parameters", ""},
            {"Parameter", "Value"},
            {"BaseDate", baseDate.ToOADate()},
            {"RateIndexName", RateIndices.JIBAR.ToString()},
            {"RateIndexTenor", "3m"},
            {"Interpolation", CurveInterpolationMethods.Exponential_On_DiscountFactors.ToString()},
        };

        Dictionary<string, double> depositRates = 
            new()
            {
                ["1m"] = 0.1,
                ["3m"] = 0.1,
                ["6m"] = 0.1,
            };
        
        object[,] depositInstruments = 
        {
            {"Deposits", "", "", ""},
            {"Tenors", "RateIndex", "Rates", "Include"},
            {"1m", RateIndices.JIBAR.ToString(), depositRates["1m"], "TRUE"},
            {"3m", RateIndices.JIBAR.ToString(), depositRates["3m"], "TRUE"},
            {"6m", RateIndices.JIBAR.ToString(), depositRates["6m"], "TRUE"},
        };
       
        Dictionary<string, double> fraRates = 
            new()
            {
                ["6x9"] = 0.1,
                ["9x12"] = 0.1,
            };
        
        object[,] fraInstruments = 
        {
            {"FRAs", "", "", ""},
            {"FraTenors", "RateIndex", "Rates", "Include"},
            {"6x9", RateIndices.JIBAR.ToString(), fraRates["6x9"], "TRUE"},
            {"9x12", RateIndices.JIBAR.ToString(), fraRates["9x12"], "TRUE"},
        };

        object[] instruments = {depositInstruments, fraInstruments};
        QL.Jibar jibar = new(new QL.Period("3m"));
        QL.DayCounter dayCounter = jibar.dayCounter();
        string handle = CurveBootstrapper.Bootstrap("BootstrappedSingleCurve", curveParameters, null, instruments);
        QL.YieldTermStructure? curve = CurveUtils.GetCurveObject(handle);
        const double tolerance = 1e-6;

        Debug.Assert(curve != null, nameof(curve) + " != null");
        Assert.AreEqual(1.0, curve.discount(baseDate.ToQuantLibDate()));
        
        DateTime date1M = (DateTime) DateUtils.AddTenorToDate(baseDate, "1m", "ZAR", "ModFol");
        double discountFactor1M = 1 / (1 + depositRates["1m"] * dayCounter.yearFraction(baseDate.ToQuantLibDate(), date1M.ToQuantLibDate()));
        Assert.AreEqual(discountFactor1M, curve.discount(date1M.ToQuantLibDate()), tolerance);
        
        DateTime date3M = (DateTime) DateUtils.AddTenorToDate(baseDate, "3m", "ZAR", "ModFol");
        double discountFactor3M = 
            1 / (1 + depositRates["3m"] * dayCounter.yearFraction(baseDate.ToQuantLibDate(), date3M.ToQuantLibDate()));
        
        Assert.AreEqual(discountFactor3M, curve.discount(date3M.ToQuantLibDate()), tolerance);
        
        DateTime date6M = (DateTime) DateUtils.AddTenorToDate(baseDate, "6m", "ZAR", "ModFol");
        double discountFactor6M = 
            1 / (1 + depositRates["6m"] * dayCounter.yearFraction(baseDate.ToQuantLibDate(), date6M.ToQuantLibDate()));
        
        Assert.AreEqual(discountFactor6M, curve.discount(date6M.ToQuantLibDate()), tolerance); 
       
        DateTime date9M = (DateTime) DateUtils.AddTenorToDate(baseDate, "9m", "ZAR", "ModFol");
        double discountFactor9M = 
            discountFactor6M * 
            1 / (1 + fraRates["6x9"] * dayCounter.yearFraction(date6M.ToQuantLibDate(), date9M.ToQuantLibDate()));
        
        Assert.AreEqual(discountFactor9M, curve.discount(date9M.ToQuantLibDate()), tolerance); 
        
        DateTime date12M = (DateTime) DateUtils.AddTenorToDate(baseDate, "12m", "ZAR", "ModFol");
        double discountFactor12M = 
            discountFactor9M * 
            1 / (1 + fraRates["9x12"] * dayCounter.yearFraction(date9M.ToQuantLibDate(), date12M.ToQuantLibDate()));
        
        Assert.AreEqual(discountFactor12M, curve.discount(date12M.ToQuantLibDate()), tolerance); 
    }
    
    [Test]
    public void Bootstrap_FlatCurve_DepositsFrasAndSwaps_Test()
    {
        DateTime baseDate = new(2022, 06, 01);

        object[,] curveParameters =
        {
            {"CurveUtils Parameters", ""},
            {"Parameter", "Value"},
            {"BaseDate", baseDate.ToOADate()},
            {"RateIndexName", RateIndices.JIBAR.ToString()},
            {"RateIndexTenor", "3m"},
            {"Interpolation", CurveInterpolationMethods.Exponential_On_DiscountFactors.ToString()},
        };

        Dictionary<string, double> depositRates = 
            new()
            {
                ["1m"] = 0.1,
                ["3m"] = 0.1,
                ["6m"] = 0.1,
            };
        
        object[,] depositInstruments = 
        {
            {"Deposits", "", "", ""},
            {"Tenors", "RateIndex", "Rates", "Include"},
            {"1m", RateIndices.JIBAR.ToString(), depositRates["1m"], "TRUE"},
            {"3m", RateIndices.JIBAR.ToString(), depositRates["3m"], "TRUE"},
            {"6m", RateIndices.JIBAR.ToString(), depositRates["6m"], "TRUE"},
        };
       
        Dictionary<string, double> fraRates = 
            new()
            {
                ["6x9"] = 0.1,
                ["9x12"] = 0.1,
            };
        
        object[,] fraInstruments = 
        {
            {"FRAs", "", "", ""},
            {"FraTenors", "RateIndex", "Rates", "Include"},
            {"6x9", RateIndices.JIBAR.ToString(), fraRates["6x9"], "TRUE"},
            {"9x12", RateIndices.JIBAR.ToString(), fraRates["9x12"], "TRUE"},
        };

        Dictionary<string, double> swapRates = 
            new()
            {
                ["2y"] = 0.1,
                ["3y"] = 0.1,
            };
        
        object[,] swapInstruments =
        {
            {"Interest Rate Swaps", "", "", ""},
            {"Tenors", "RateIndex", "Rates", "Include"},
            {"2y", RateIndices.JIBAR.ToString(), swapRates["2y"], "TRUE"},
            {"3y", RateIndices.JIBAR.ToString(), swapRates["3y"], "TRUE"},
        };

        object[] instruments = {depositInstruments, fraInstruments, swapInstruments};
        QL.Actual365Fixed dayCounter = new();
        string handle = 
            CurveBootstrapper.Bootstrap(
                handle: nameof(Bootstrap_FlatCurve_DepositsFrasAndSwaps_Test), 
                curveParameters: curveParameters, 
                customRateIndex: null, 
                instrumentGroups: instruments);
        
        QL.YieldTermStructure? curve = CurveUtils.GetCurveObject(handle);
        const double tolerance = 0.0001;
        List<string> tenors = new() {"0m", "1m", "3m", "6m", "9m", "12m", "15m", "18m", "21m", "24m"};
        Dictionary<string, DateTime> dates = tenors.ToDictionary(t => t, t => (DateTime)DateUtils.AddTenorToDate(baseDate, t, "ZAR", "ModFol"));
        
        double discountFactor1M = 
            1 / 
            (1 + depositRates["1m"] * dayCounter.yearFraction(baseDate.ToQuantLibDate(), dates["1m"].ToQuantLibDate()));

        Debug.Assert(curve != null, nameof(curve) + " != null");
        Assert.AreEqual(discountFactor1M, curve.discount(dates["1m"].ToQuantLibDate()), tolerance);
        
        double discountFactor3M = 
            1 / 
            (1 + depositRates["3m"] * dayCounter.yearFraction(baseDate.ToQuantLibDate(), dates["3m"].ToQuantLibDate()));
        
        Assert.AreEqual(discountFactor3M, curve.discount(dates["3m"].ToQuantLibDate()), tolerance);
        
        double discountFactor6M = 
            1 / 
            (1 + depositRates["6m"] * dayCounter.yearFraction(baseDate.ToQuantLibDate(), dates["6m"].ToQuantLibDate()));
        
        Assert.AreEqual(discountFactor6M, curve.discount(dates["6m"].ToQuantLibDate()), tolerance); 
       
        double discountFactor9M = 
            discountFactor6M / 
            (1 + fraRates["6x9"] * dayCounter.yearFraction(dates["6m"].ToQuantLibDate(), dates["9m"].ToQuantLibDate()));
        
        Assert.AreEqual(discountFactor9M, curve.discount(dates["9m"].ToQuantLibDate()), tolerance); 
        
        double discountFactor12M = 
            discountFactor9M / 
            (1 + fraRates["9x12"] * dayCounter.yearFraction(dates["9m"].ToQuantLibDate(), dates["12m"].ToQuantLibDate()));

        Assert.AreEqual(discountFactor12M, curve.discount(dates["12m"].ToQuantLibDate()), tolerance);

        List<double> startDatesRange = 
            dates.Values.ToList().GetRange(0, dates.Count - 1).Select(d => d.ToOADate()).ToList();
        
        List<double> endDatesRange = 
            dates.Values.ToList().GetRange(1, dates.Count - 1).Select(d => d.ToOADate()).ToList();
        
        List<double> forwardRates = 
            ExcelArrayUtils.ConvertExcelRangeToList<double>(
                (object[,])CurveUtils.GetForwardRates(
                    handle: handle, 
                    startDatesRange: ExcelArrayUtils.ConvertListToExcelRange(startDatesRange, 0), 
                    endDatesRange: ExcelArrayUtils.ConvertListToExcelRange(endDatesRange, 0), 
                    compoundingConventionParameter: "Simple"));

       List<double> discountFactors =
            ExcelArrayUtils.ConvertExcelRangeToList<double>(
                (object[,])CurveUtils.GetDiscountFactors(handle, endDatesRange.Cast<object>().ToArray()));
       
       List<double> dayCountFractions = 
           endDatesRange
               .Zip(
                   startDatesRange, 
                   (s, e) => 
                       dayCounter.yearFraction(
                           d1: DateTime.FromOADate(s).ToQuantLibDate(), 
                           d2: DateTime.FromOADate(e).ToQuantLibDate()))
               .ToList();
       
       double numerator =
           forwardRates.Zip(
                dayCountFractions.Zip(
                    discountFactors, (t, df) => t * df), 
            (f, tDf) => f * tDf).Sum(); 
       
       double denominator =
                dayCountFractions.Zip(
                    discountFactors, (t, df) => t * df).Sum(); 
       
       double parSwapRate2Y = numerator / denominator;
    
       Assert.AreEqual(swapRates["2y"], parSwapRate2Y, tolerance);
    }
    
    [Test]
    public void Bootstrap_BackwardFlatInterpolation_Test()
    {
        QL.Date baseDate = new(1, 6.ToQuantLibMonth(), 2022);
            
        object[,] curveParameters =
        {
            {"CurveUtils Parameters", ""},
            {"Parameter", "Value"},
            {"BaseDate", baseDate.ToDateTime().ToOADate()},
            {"RateIndexName", RateIndices.JIBAR.ToString()},
            {"RateIndexTenor", "3m"},
            {"Interpolation", CurveInterpolationMethods.Exponential_On_DiscountFactors.ToString()},
        };
            
        object[,] depositInstruments = 
        {
            {"Deposits", "", "", ""},
            {"Tenors", "RateIndex", "Rates", "Include"},
            {"3m", RateIndices.JIBAR.ToString(), 0.1, "TRUE"},
            {"6m", RateIndices.JIBAR.ToString(), 0.1, "TRUE"},
        };
                   
        object[,] fraInstruments = 
        {
            {"FRAs", "", "", ""},
            {"FraTenors", "RateIndex", "Rates", "Include"},
            {"6x9", RateIndices.JIBAR.ToString(), 0.1, "TRUE"},
            {"9x12", RateIndices.JIBAR.ToString(), 0.1, "TRUE"},
        };
            
        object[,] swapInstruments =
        {
            {"Interest Rate Swaps", "", "", ""},
            {"Tenors", "RateIndex", "Rates", "Include"},
            {"2y", RateIndices.JIBAR.ToString(), 0.1, "TRUE"},
            {"3y", RateIndices.JIBAR.ToString(), 0.1, "TRUE"},
        };
            
        object[] instruments = {depositInstruments, fraInstruments, swapInstruments};
        QL.Jibar jibar = new(new QL.Period(3, QL.TimeUnit.Months));
        QL.DayCounter? dayCounter = jibar.dayCounter();
        string handle = 
            CurveBootstrapper.Bootstrap(
                handle: nameof(Bootstrap_BackwardFlatInterpolation_Test), 
                curveParameters: curveParameters, 
                customRateIndex: null, 
                instrumentGroups: instruments);
        
        QL.YieldTermStructure? curve = CurveUtils.GetCurveObject(handle);
        const double tolerance = 0.000000001;
        QL.Date date3M = ((DateTime)DateUtils.AddTenorToDate(baseDate.ToDateTime(), "3M", "ZAR", "ModFol")).ToQuantLibDate();
        Debug.Assert(curve != null, nameof(curve) + " != null");
        Assert.AreEqual(1.0, curve.discount(baseDate));
        Assert.AreEqual(
            expected: 1 / (1 + 0.1 * dayCounter.yearFraction(baseDate, date3M)),
            actual: curve.discount(date3M),
            delta: tolerance);
        
        // Assert.AreEqual(
        //     expected: Math.Exp(-0.105 * dayCounter.yearFraction(baseDate, baseDate.AddMonths(3))),
        //     actual: curve.discount(baseDate.AddMonths(3)),
        //     delta: tolerance);
        // Assert.AreEqual(
        //     expected: Math.Exp(-0.13 * dayCounter.yearFraction(baseDate, baseDate.AddMonths(5))),
        //     actual: curve.discount(baseDate.AddMonths(5)),
        //     delta: tolerance);

        QL.Calendar calendar = jibar.fixingCalendar();

        double discountFactor1Y =
            Math.Exp(-0.130 * dayCounter.yearFraction(
                d1: baseDate, 
                d2: calendar.advance(baseDate, new QL.Period(6, QL.TimeUnit.Months)))) *
            Math.Exp(-0.135 * dayCounter.yearFraction(
                d1: calendar.advance(baseDate, new QL.Period(6, QL.TimeUnit.Months)), 
                d2: calendar.advance(baseDate, new QL.Period(9, QL.TimeUnit.Months))));

        int tenor = 9;
        double expectedNaccZeroRate1Y =
            -1 * Math.Log(discountFactor1Y) / 
            dayCounter.yearFraction(baseDate, baseDate.ToDateTime().AddMonths(tenor).ToQuantLibDate());
        double actualNaccZeroRate1Y = 
            -1 * Math.Log(curve.discount(calendar.advance(baseDate, new QL.Period(tenor, QL.TimeUnit.Months)))) / 
            dayCounter.yearFraction(baseDate, calendar.advance(baseDate, new QL.Period(tenor, QL.TimeUnit.Months)));
        // Assert.AreEqual(expectedNaccZeroRate1Y, actualNaccZeroRate1Y);
        
        // Assert.AreEqual(
        //     expected: discountFactor1Y,
        //     actual: curve.discount(baseDate.AddMonths(tenor)),
        //     delta: tolerance);
        // Assert.AreEqual(
        //     expected: Math.Exp(-0.1 * dayCounter.yearFraction(baseDate, baseDate.AddMonths(6))),
        //     actual: curve.discount(baseDate.AddMonths(6)),
        //     delta: tolerance);
        // Assert.AreEqual(
        //     expected: Math.Exp(-0.1 * dayCounter.yearFraction(baseDate, baseDate.AddMonths(6))) *
        //               Math.Exp(-0.1 * dayCounter.yearFraction(baseDate.AddMonths(6), baseDate.AddMonths(9))),
        //     actual: curve.discount(baseDate.AddMonths(9)),
        //     delta: tolerance);
    }

    [Test]
    public void Bootstrap_Get_ZarSwapCurveTest()
    {
        // TODO: Needs to warn user if not connected to VPN.
        DateTime baseDate = new(2023, 03, 31);
        string handle = CurveBootstrapper.Get("ZarSwapCurve", "ZAR-Swap", baseDate);
        object[] dates = {baseDate.ToOADate(), baseDate.AddYears(1).ToOADate()};
        object discountFactors = CurveUtils.GetDiscountFactors(handle, dates);
        Assert.AreEqual(((object[,])discountFactors)[0, 0], 1.000000000000d); 
        Assert.AreEqual((double)((object[,])discountFactors)[1, 0], 0.922680208459d, 1e-10d);
    }

    [Test]
    public void GetAvailableBootstrappingInterpolationMethods()
    {
        Array methods = Enum.GetValues(typeof(CurveInterpolationMethods));
        object[,] expectedOutput = new object[methods.Length + 1, 1];
        expectedOutput[0, 0] = "IR Bootstrapping Interpolation Methods";
        int i = 0;
        foreach (CurveInterpolationMethods method in methods)
        {
            expectedOutput[i++, 0] = method.ToString();
        }
        
        object[,] actualOutput = (object[,])CurveBootstrapper.GetBootstrappingInterpolationMethods();
        
        Assert.AreEqual(expectedOutput, actualOutput);
    }
}
