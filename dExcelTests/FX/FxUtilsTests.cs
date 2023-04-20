namespace dExcelTests.FX;

using dExcel.Dates;
using dExcel.ExcelUtils;
using dExcel.FX;
using NUnit.Framework;
using QLNet;
using QL = QuantLib;

[TestFixture]
public class FxUtilsTests
{
    [Test]
    [TestCase(10.0, 10.0, 0.1, 0.1, 0.25, 1.0, "CALL")]
    [TestCase(15.0, 10.0, 0.1, 0.1, 0.25, 1.0, "CALL")]
    [TestCase(10.0, 10.0, 0.1, 0.1, 0.25, 1.0, "PUT")]
    [TestCase(15.0, 10.0, 0.1, 0.1, 0.25, 1.0, "PUT")]
    public void CalculateDeltaForCallOptionTest(
        double spot, 
        double strike, 
        double domesticRate, 
        double foreignRate, 
        double vol, 
        double optionMaturity,
        string optionType)
    {
        double delta = FxUtils.CalculateDelta(spot, strike, domesticRate, foreignRate, vol, optionMaturity, optionType);
        double optionPrice = 
            (double) Pricers.GarmanKohlhagenSpotOptionPricer(
                spotPrice: spot, 
                strike: strike, 
                domesticRiskFreeRate: domesticRate, 
                foreignRiskFreeRate: foreignRate, 
                vol: vol,
                optionMaturity: optionMaturity, 
                optionType: optionType, 
                direction: "LONG");

        const double bump = 1e-5;
        double bumpedOptionPrice = 
            (double)Pricers.GarmanKohlhagenSpotOptionPricer(
                spotPrice: spot + bump, 
                strike: strike, 
                domesticRiskFreeRate: domesticRate, 
                foreignRiskFreeRate: foreignRate, 
                vol: vol, 
                optionMaturity: optionMaturity, 
                optionType, 
                direction: "LONG");
    
        Assert.AreEqual(delta, (bumpedOptionPrice - optionPrice)/bump, 1e-5);
    }

    [Test]
    public void ConvertDeltaToMoneynessVolSurfaceTest()
    {
        // Note both the moneyness and delta based vol surfaces were extracted from Eikon for 2023-03-31.
        // EURUSD vol surface.
        DateTime baseDate = new(2023, 03, 31);
        
        double[] moneynesses =
            {-10.00, -8.75, -7.50, -6.25, -5.00, -3.75, -2.50, -1.25, 0.00, 1.25, 2.5, 3.75, 5, 6.25, 7.5, 8.75, 10};
        
        List<string> tenors = 
            new() {"ON", "SW", "2W", "1M", "2M", "3M", "6M", "9M", "1Y", "2Y", "3Y", "5Y", "7Y", "10Y"};

        object[,] optionMaturities =
            ExcelArrayUtils.ConvertListToExcelRange(
                tenors
                    .Select(t =>
                        DateUtils.Act360(baseDate, (DateTime)DateUtils.AddTenorToDate(baseDate, t, "EUR", "ModFol")))
                    .ToList(), 0);
        
        double[,] moneynessVols =
        {
            { 0.14430, 0.13551, 0.12613, 0.11603, 0.10505, 0.09295, 0.07945, 0.06462, 0.05297, 0.06254, 0.08298, 0.10157, 0.11779, 0.13222, 0.14532, 0.15737, 0.16859 },
            { 0.13432, 0.12765, 0.12069, 0.11346, 0.10596, 0.09828, 0.09070, 0.08394, 0.07966, 0.08043, 0.08704, 0.09706, 0.10806, 0.11894, 0.12937, 0.13926, 0.14864 },
            { 0.11616, 0.11242, 0.10856, 0.10458, 0.10045, 0.09620, 0.09185, 0.08759, 0.08430, 0.08398, 0.08568, 0.08786, 0.09015, 0.09244, 0.09471, 0.09695, 0.09915 },
            { 0.10591, 0.10220, 0.09838, 0.09447, 0.09047, 0.08644, 0.08251, 0.07901, 0.07655, 0.07581, 0.07667, 0.07843, 0.08061, 0.08296, 0.08537, 0.08779, 0.09020 },
            { 0.10384, 0.10073, 0.09755, 0.09431, 0.09104, 0.08777, 0.08462, 0.08180, 0.07970, 0.07870, 0.07880, 0.07961, 0.08081, 0.08220, 0.08370, 0.08524, 0.08680 },
            { 0.10342, 0.10016, 0.09689, 0.09361, 0.09038, 0.08728, 0.08443, 0.08197, 0.08013, 0.07903, 0.07874, 0.07915, 0.08010, 0.08143, 0.08302, 0.08476, 0.08660 },
            { 0.09990, 0.09712, 0.09434, 0.09158, 0.08887, 0.08627, 0.08385, 0.08173, 0.08000, 0.07878, 0.07811, 0.07799, 0.07834, 0.07907, 0.08007, 0.08128, 0.08261 },
            { 0.09722, 0.09482, 0.09242, 0.09002, 0.08765, 0.08536, 0.08322, 0.08130, 0.07970, 0.07853, 0.07782, 0.07758, 0.07771, 0.07813, 0.07877, 0.07955, 0.08044 },
            { 0.09597, 0.09367, 0.09136, 0.08908, 0.08685, 0.08471, 0.08271, 0.08091, 0.07940, 0.07824, 0.07747, 0.07708, 0.07704, 0.07728, 0.07774, 0.07837, 0.07913 },
            { 0.09097, 0.08945, 0.08798, 0.08656, 0.08521, 0.08394, 0.08277, 0.08172, 0.08080, 0.08003, 0.07942, 0.07899, 0.07872, 0.07862, 0.07869, 0.07890, 0.07925 },
            { 0.08857, 0.08755, 0.08657, 0.08564, 0.08478, 0.08399, 0.08327, 0.08264, 0.08210, 0.08166, 0.08132, 0.08108, 0.08095, 0.08093, 0.08101, 0.08119, 0.08147 },
            { 0.09083, 0.09019, 0.08958, 0.08902, 0.08849, 0.08802, 0.08759, 0.08722, 0.08690, 0.08664, 0.08643, 0.08629, 0.08620, 0.08618, 0.08621, 0.08630, 0.08645 },
            { 0.09331, 0.09284, 0.09240, 0.09199, 0.09162, 0.09128, 0.09098, 0.09072, 0.09050, 0.09032, 0.09018, 0.09009, 0.09004, 0.09002, 0.09005, 0.09013, 0.09024 },
            { 0.09594, 0.09558, 0.09525, 0.09493, 0.09465, 0.09440, 0.09417, 0.09397, 0.09380, 0.09366, 0.09355, 0.09347, 0.09342, 0.09340, 0.09341, 0.09345, 0.09352 },
        };

        object[,] deltas =
            {{ 0.90, 0.85, 0.80, 0.75, 0.70, 0.65, 0.60, 0.55, 0.50, 0.45, 0.40, 0.35, 0.30, 0.25, 0.20, 0.15, 0.10 }};

        object[,] deltaVols =
        {
            { 0.05805, 0.05662, 0.05564, 0.05490, 0.05432, 0.05386, 0.05350, 0.05320, 0.05297, 0.05281, 0.05270, 0.05267, 0.05272, 0.05287, 0.05317, 0.05370, 0.05470 },
            { 0.08522, 0.08370, 0.08265, 0.08184, 0.08121, 0.08069, 0.08027, 0.07993, 0.07966, 0.07946, 0.07931, 0.07924, 0.07925, 0.07937, 0.07964, 0.08014, 0.08109 },
            { 0.09110, 0.08947, 0.08825, 0.08727, 0.08645, 0.08576, 0.08517, 0.08468, 0.08430, 0.08401, 0.08382, 0.08373, 0.08375, 0.08387, 0.08411, 0.08450, 0.08510 },
            { 0.08470, 0.08259, 0.08107, 0.07988, 0.07893, 0.07814, 0.07750, 0.07697, 0.07655, 0.07623, 0.07599, 0.07585, 0.07581, 0.07588, 0.07610, 0.07650, 0.07720 },
            { 0.09037, 0.08770, 0.08574, 0.08419, 0.08292, 0.08187, 0.08099, 0.08026, 0.07970, 0.07926, 0.07892, 0.07871, 0.07863, 0.07869, 0.07890, 0.07931, 0.08002 },
            { 0.09295, 0.08954, 0.08712, 0.08526, 0.08378, 0.08257, 0.08157, 0.08075, 0.08013, 0.07962, 0.07920, 0.07891, 0.07875, 0.07875, 0.07895, 0.07940, 0.08030 },
            { 0.09703, 0.09257, 0.08937, 0.08690, 0.08490, 0.08327, 0.08191, 0.08078, 0.08000, 0.07935, 0.07873, 0.07829, 0.07803, 0.07798, 0.07819, 0.07874, 0.07984 },
            { 0.09866, 0.09383, 0.09030, 0.08752, 0.08525, 0.08336, 0.08179, 0.08049, 0.07970, 0.07904, 0.07833, 0.07785, 0.07761, 0.07760, 0.07787, 0.07847, 0.07957 },
            { 0.10062, 0.09510, 0.09109, 0.08795, 0.08541, 0.08331, 0.08158, 0.08015, 0.07940, 0.07875, 0.07794, 0.07738, 0.07707, 0.07703, 0.07729, 0.07793, 0.07914 },
            { 0.10005, 0.09469, 0.09090, 0.08800, 0.08570, 0.08383, 0.08231, 0.08108, 0.08080, 0.08054, 0.07969, 0.07907, 0.07871, 0.07863, 0.07890, 0.07966, 0.08124 },
            { 0.09799, 0.09324, 0.08994, 0.08748, 0.08558, 0.08408, 0.08291, 0.08203, 0.08210, 0.08218, 0.08150, 0.08107, 0.08092, 0.08109, 0.08166, 0.08284, 0.08505 },
            { 0.10098, 0.09645, 0.09334, 0.09106, 0.08935, 0.08806, 0.08712, 0.08651, 0.08690, 0.08745, 0.08671, 0.08628, 0.08619, 0.08646, 0.08720, 0.08861, 0.09117 },
            { 0.10459, 0.09983, 0.09663, 0.09434, 0.09266, 0.09144, 0.09062, 0.09015, 0.09050, 0.09108, 0.09039, 0.09006, 0.09009, 0.09053, 0.09150, 0.09324, 0.09630 },
            { 0.10792, 0.10273, 0.09934, 0.09698, 0.09533, 0.09422, 0.09359, 0.09341, 0.09380, 0.09513, 0.09410, 0.09353, 0.09342, 0.09381, 0.09483, 0.09676, 0.10026 },
        };
        
        // object convertedVolSurface = FxUtils.ConvertDeltaToMoneynessVolSurface(deltaVols, deltas, optionMaturities, 0.1, 0.2);
    }


    [Test]
    public void InterpolationTest()
    {
        Cubic interpolationFactory = new();
        List<double> xBegin = new() { 0.0, 1.0, 2.0, 3.0, 4.0, 5.0};
        List<double> yBegin = new() { 0.0, 1.0, 2.0, 3.0, 4.0, 5.0};
        CubicNaturalSpline splineNatural = new(xBegin, xBegin.Count, yBegin);
        Discount discount = new();
        discount.discountImpl(splineNatural, 0.1);
        // InterpolatedDiscountCurve<IInterpolationFactory> test = 
        // new InterpolatedDiscountCurve<IInterpolationFactory>(new Actual360(), null, null, splineNatural);
        QL.DateVector usdCurveDates = 
            new(
                new List<QL.Date>()
                {
                    new(31, QL.Month.March, 2023),
                    new(3, QL.Month.April, 2023),
                    new(4, QL.Month.April, 2023),
                    new(11, QL.Month.April, 2023),
                    new(4, QL.Month.May, 2023),
                    new(5, QL.Month.June, 2023),
                    new(5, QL.Month.July, 2023),
                    new(4, QL.Month.October, 2023),
                    new(4, QL.Month.January, 2024),
                    new(4, QL.Month.April, 2024),
                    new(4, QL.Month.April, 2025),
                    new(6, QL.Month.April, 2026),
                    new(5, QL.Month.April, 2027),
                    new(4, QL.Month.April, 2028),
                    new(4, QL.Month.April, 2029),
                    new(4, QL.Month.April, 2030),
                    new(4, QL.Month.April, 2031),
                    new(5, QL.Month.April, 2032),
                    new(4, QL.Month.April, 2033),
                    new(4, QL.Month.April, 2035),
                    new(5, QL.Month.April, 2038),
                    new(6, QL.Month.April, 2043),
                    new(6, QL.Month.April, 2048),
                    new(4, QL.Month.April, 2053),
                    new(4, QL.Month.April, 2063),
                    new(4, QL.Month.April, 2073),
                });

        QL.DoubleVector usdCurveValues = new(new List<double>()
        {
            1, 0.999594748, 0.999459682, 0.998485477, 0.995236637, 0.99065656, 0.986389473, 0.973558139, 0.961709557,
            0.951187072, 0.917197683, 0.888698879, 0.86212352, 0.836076594, 0.810267688, 0.785011688, 0.760210466,
            0.735227273, 0.710992469, 0.663886482, 0.5985451, 0.509908045, 0.446464292, 0.395983824, 0.331760479,
            0.293876051
        });
        
        QL.DiscountCurve usdDiscountCurve = new QL.DiscountCurve(usdCurveDates, usdCurveValues, new QL.Actual360());
        
    }
}
