namespace dExcelTests.InterestRates;

using dExcel.Dates;
using dExcel.InterestRates;
using NUnit.Framework;
using QLNet;

[TestFixture]
public class SingleCurveBootstrapperTest
{
    [Test]
    public void BootstrapTest()
    {
        object[,] curveParameters = 
        {
            {"Parameter", "Value"},
            {"BaseDate", 44713},
            {"RateIndexName", "JIBAR"},
            {"Interpolation", "Exponential"},
        };

        object[,] deposits =
        {
            {"Tenors", "Rates", "Include"},
            {"1m", "10%", "TRUE"},
            {"2m", "10%", "TRUE"},
            {"3m", "10%", "TRUE"},
        };

        object[,] fras =
        {
            {"FraTenors", "Rates", "Include"},
            {"3x6", "10%", "TRUE"},
            {"6x9", "10%", "TRUE"},
            {"9x12", "10%", "TRUE"},
        };

        object[,] swaps =
        {
            {"Tenors", "Rates", "Include"},
            {"2y", "10%", "TRUE"},
            {"3y", "10%", "TRUE"},
        };
        
        object[] instruments = {deposits, fras, swaps};
        string handle = SingleCurveBootstrapper.Bootstrap("ZAR-Swap", curveParameters, null, instruments);
    }
    
    
    [Test]
    public void Bootstrap_MissingBaseDate_Test()
    {
        object[,] curveParameters =
        {
            {"CurveUtils Parameters", ""},
            {"Parameter", "Value"},
            {"RateIndexName", "JIBAR"},
            {"RateIndexTenor", "3m"},
            {"Interpolation", "Linear"},
        };

        object[,] instrumentGroups = 
        {
            {"Deposits", "", "", ""},
            {"Tenors", "RateIndex", "Rates", "Include"},
            {"1m", "JIBAR", 0.1, "TRUE"},
            {"3m", "JIBAR", 0.1, "TRUE"},
            {"6m", "JIBAR", 0.1, "TRUE"},
        };
        
        string handle = 
            SingleCurveBootstrapper.Bootstrap(
                handle: "BootstrappedSingleCurve", 
                curveParameters: curveParameters,
                customRateIndex: null,
                instrumentGroups: instrumentGroups);
        
            const string expected = "#∂Excel Error: Curve parameter missing: 'BASEDATE'.";
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
            {"Interpolation", "Linear"},
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
            {"1m", "JIBAR", depositRates["1m"], "TRUE"},
            {"3m", "JIBAR", depositRates["3m"], "TRUE"},
            {"6m", "JIBAR", depositRates["6m"], "TRUE"},
        };
        
        DayCounter dayCounter = new Actual365Fixed();
        string handle = 
            SingleCurveBootstrapper.Bootstrap(
                handle: "BootstrappedSingleCurve", 
                curveParameters, 
                customRateIndex: null,
                instruments);

        YieldTermStructure? curve = CurveUtils.GetCurveObject(handle);
        const double tolerance = 0.01; 
        Assert.AreEqual(1.0, curve.discount(baseDate));
                    
        DateTime date1M = (DateTime) DateUtils.AddTenorToDate(baseDate, "1m", "ZAR", "ModFol");
        double discountFactor1M = 1 / (1 + depositRates["1m"] * dayCounter.yearFraction(baseDate, date1M));
        Assert.AreEqual(discountFactor1M, curve.discount(date1M), tolerance);
        
        DateTime date3M = (DateTime) DateUtils.AddTenorToDate(baseDate, "3m", "ZAR", "ModFol");
        double discountFactor3M = 1 / (1 + depositRates["3m"] * dayCounter.yearFraction(baseDate, date3M));
        Assert.AreEqual(discountFactor3M, curve.discount(date3M), tolerance);
        
        DateTime date6M = (DateTime) DateUtils.AddTenorToDate(baseDate, "6m", "ZAR", "ModFol");
        double discountFactor6M = 1 / (1 + depositRates["6m"] * dayCounter.yearFraction(baseDate, date6M));
        Assert.AreEqual(discountFactor6M, curve.discount(date6M), tolerance); 
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
            {"RateIndexName", "JIBAR"},
            {"RateIndexTenor", "3m"},
            {"Interpolation", "Linear"},
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
            {"1m", "JIBAR", depositRates["1m"], "TRUE"},
            {"3m", "JIBAR", depositRates["3m"], "TRUE"},
            {"6m", "JIBAR", depositRates["6m"], "TRUE"},
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
            {"6x9", "JIBAR", fraRates["6x9"], "TRUE"},
            {"9x12", "JIBAR", fraRates["9x12"], "TRUE"},
        };

        object[] instruments = {depositInstruments, fraInstruments};
        Jibar jibar = new(new Period("3m"));
        DayCounter dayCounter = jibar.dayCounter();
        string handle = SingleCurveBootstrapper.Bootstrap("BootstrappedSingleCurve", curveParameters, null, instruments);
        YieldTermStructure curve = CurveUtils.GetCurveObject(handle);
        const double tolerance = 1e-6; 
        
        Assert.AreEqual(1.0, curve.discount(baseDate));
        
        DateTime date1M = (DateTime) DateUtils.AddTenorToDate(baseDate, "1m", "ZAR", "ModFol");
        double discountFactor1M = 1 / (1 + depositRates["1m"] * dayCounter.yearFraction(baseDate, date1M));
        Assert.AreEqual(discountFactor1M, curve.discount(date1M), tolerance);
        
        DateTime date3M = (DateTime) DateUtils.AddTenorToDate(baseDate, "3m", "ZAR", "ModFol");
        double discountFactor3M = 1 / (1 + depositRates["3m"] * dayCounter.yearFraction(baseDate, date3M));
        Assert.AreEqual(discountFactor3M, curve.discount(date3M), tolerance);
        
        DateTime date6M = (DateTime) DateUtils.AddTenorToDate(baseDate, "6m", "ZAR", "ModFol");
        double discountFactor6M = 1 / (1 + depositRates["6m"] * dayCounter.yearFraction(baseDate, date6M));
        Assert.AreEqual(discountFactor6M, curve.discount(date6M), tolerance); 
       
        DateTime date9M = (DateTime) DateUtils.AddTenorToDate(baseDate, "9m", "ZAR", "ModFol");
        double discountFactor9M = 
            discountFactor6M * 1 / (1 + fraRates["6x9"] * dayCounter.yearFraction(date6M, date9M));
        
        Assert.AreEqual(discountFactor9M, curve.discount(date9M), tolerance); 
        
        DateTime date12M = (DateTime) DateUtils.AddTenorToDate(baseDate, "12m", "ZAR", "ModFol");
        double discountFactor12M = 
            discountFactor9M * 1 / (1 + fraRates["9x12"] * dayCounter.yearFraction(date9M, date12M));
        
        Assert.AreEqual(discountFactor12M, curve.discount(date12M), tolerance); 
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
            {"RateIndexName", "JIBAR"},
            {"RateIndexTenor", "3m"},
            {"Interpolation", "Linear"},
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
            {"1m", "JIBAR", depositRates["1m"], "TRUE"},
            {"3m", "JIBAR", depositRates["3m"], "TRUE"},
            {"6m", "JIBAR", depositRates["6m"], "TRUE"},
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
            {"6x9", "JIBAR", fraRates["6x9"], "TRUE"},
            {"9x12", "JIBAR", fraRates["9x12"], "TRUE"},
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
            {"2y", "JIBAR", swapRates["2y"], "TRUE"},
            {"3y", "JIBAR", swapRates["3y"], "TRUE"},
        };

        object[] instruments = {depositInstruments, fraInstruments, swapInstruments};
        Actual365Fixed dayCounter = new();
        string handle = SingleCurveBootstrapper.Bootstrap("BootstrappedSingleCurve", curveParameters, null, instruments);
        YieldTermStructure? curve = CurveUtils.GetCurveObject(handle);
        const double tolerance = 0.0001; 
        
        DateTime date1M = (DateTime) DateUtils.AddTenorToDate(baseDate, "1m", "ZAR", "ModFol");
        double discountFactor1M = 1 / (1 + depositRates["1m"] * dayCounter.yearFraction(baseDate, date1M));
        Assert.AreEqual(discountFactor1M, curve.discount(date1M), tolerance);
        
        DateTime date3M = (DateTime) DateUtils.AddTenorToDate(baseDate, "3m", "ZAR", "ModFol");
        double discountFactor3M = 1 / (1 + depositRates["3m"] * dayCounter.yearFraction(baseDate, date3M));
        Assert.AreEqual(discountFactor3M, curve.discount(date3M), tolerance);
        
        DateTime date6M = (DateTime) DateUtils.AddTenorToDate(baseDate, "6m", "ZAR", "ModFol");
        double discountFactor6M = 1 / (1 + depositRates["6m"] * dayCounter.yearFraction(baseDate, date6M));
        Assert.AreEqual(discountFactor6M, curve.discount(date6M), tolerance); 
       
        DateTime date9M = (DateTime) DateUtils.AddTenorToDate(baseDate, "9m", "ZAR", "ModFol");
        double discountFactor9M = 
            discountFactor6M * 1 / (1 + fraRates["6x9"] * dayCounter.yearFraction(date6M, date9M));
        
        Assert.AreEqual(discountFactor9M, curve.discount(date9M), tolerance); 
        
        DateTime date12M = (DateTime) DateUtils.AddTenorToDate(baseDate, "12m", "ZAR", "ModFol");
        double discountFactor12M = 
            discountFactor9M * 1 / (1 + fraRates["9x12"] * dayCounter.yearFraction(date9M, date12M));
        
        Assert.AreEqual(discountFactor12M, curve.discount(date12M), tolerance); 
    }


    [Test]
    public void Bootstrap_BackwardFlatInterpolation_Test()
    {
        DateTime baseDate = new(2022, 06, 01);
            
        object[,] curveParameters =
        {
            {"CurveUtils Parameters", ""},
            {"Parameter", "Value"},
            {"BaseDate", baseDate.ToOADate()},
            {"RateIndexName", "JIBAR"},
            {"RateIndexTenor", "3m"},
            {"Interpolation", "Linear"},
        };
            
        object[,] depositInstruments = 
        {
            {"Deposits", "", "", ""},
            {"Tenors", "RateIndex", "Rates", "Include"},
            {"3m", "JIBAR", 0.1, "TRUE"},
            {"6m", "JIBAR", 0.1, "TRUE"},
        };
                   
            // {"1m", "JIBAR", 0.10, "TRUE"},
        object[,] fraInstruments = 
        {
            {"FRAs", "", "", ""},
            {"FraTenors", "RateIndex", "Rates", "Include"},
            {"6x9", "JIBAR", 0.1, "TRUE"},
            {"9x12", "JIBAR", 0.1, "TRUE"},
        };
            
        object[,] swapInstruments =
        {
            {"Interest Rate Swaps", "", "", ""},
            {"Tenors", "RateIndex", "Rates", "Include"},
            {"2y", "JIBAR", 0.1, "TRUE"},
            {"3y", "JIBAR", 0.1, "TRUE"},
        };
            
        object[] instruments = {depositInstruments, fraInstruments, swapInstruments};
        var jibar = new Jibar(new Period("3m"));
        var dayCounter = jibar.dayCounter();
        string handle = SingleCurveBootstrapper.Bootstrap("BootstrappedSingleCurve", curveParameters, null, instruments);
        YieldTermStructure? curve = CurveUtils.GetCurveObject(handle);
        const double tolerance = 0.000000001;
        DateTime date3m = (DateTime)DateUtils.AddTenorToDate(baseDate, "3M", "ZAR", "ModFol");
        Assert.AreEqual(1.0, curve.discount(baseDate));
        Assert.AreEqual(
            expected: 1 / (1 + 0.1 * dayCounter.yearFraction(baseDate, date3m)),
            actual: curve.discount(date3m),
            delta: tolerance);
        
        // Assert.AreEqual(
        //     expected: Math.Exp(-0.105 * dayCounter.yearFraction(baseDate, baseDate.AddMonths(3))),
        //     actual: curve.discount(baseDate.AddMonths(3)),
        //     delta: tolerance);
        // Assert.AreEqual(
        //     expected: Math.Exp(-0.13 * dayCounter.yearFraction(baseDate, baseDate.AddMonths(5))),
        //     actual: curve.discount(baseDate.AddMonths(5)),
        //     delta: tolerance);


        double discountFactor1Y =
            Math.Exp(-0.130 * dayCounter.yearFraction(baseDate, baseDate.AddMonths(6))) *
            Math.Exp(-0.135 * dayCounter.yearFraction(baseDate.AddMonths(6), baseDate.AddMonths(9)));

        int tenor = 9;
        double expectedNaccZeroRate1Y =
            -1 * Math.Log(discountFactor1Y) / dayCounter.yearFraction(baseDate, baseDate.AddMonths(tenor));
        double actualNaccZeroRate1Y = -1 * Math.Log(curve.discount(baseDate.AddMonths(tenor))) /
                                                    dayCounter.yearFraction(baseDate, baseDate.AddMonths(tenor));
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
}
