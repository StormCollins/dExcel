using dExcel.Dates;
using dExcel.InterestRates;
using NUnit.Framework;
using QL = QuantLib;

namespace dExcelTests.InterestRates;

[TestFixture]
public class FxBasisCurveBootstrapperTests
{
    [Test]
    public void BootstrapFxBasisCurveTest()
    {
        QL.Date baseDate = new(31, 3.ToQuantLibMonth(), 2023);
        
        object[,] zarForecastCurveParameters =
        {
            { "Parameter", "Value" },
            { "Calendars", "ZAR"},
            { "DayCountConvention", "Actual365" },
            { "Interpolation", "Cubic" },
        };
        
        object[,] zarForecastCurveDates = 
            {
                {new QL.Date(31, 3.ToQuantLibMonth(), 2023).ToOaDate()},
                {new QL.Date(3, 4.ToQuantLibMonth(), 2023).ToOaDate()}, 
                {new QL.Date(4, 4.ToQuantLibMonth(), 2023).ToOaDate()}, 
                {new QL.Date(11, 4.ToQuantLibMonth(), 2023).ToOaDate()}, 
                {new QL.Date(28, 4.ToQuantLibMonth(), 2023).ToOaDate()}, 
                {new QL.Date(31, 5.ToQuantLibMonth(), 2023).ToOaDate()}, 
                {new QL.Date(30, 6.ToQuantLibMonth(), 2023).ToOaDate()},
                {new QL.Date(29, 9.ToQuantLibMonth(), 2023).ToOaDate()},
                {new QL.Date(29, 12.ToQuantLibMonth(), 2023).ToOaDate()}, 
                {new QL.Date(28, 3.ToQuantLibMonth(), 2024).ToOaDate()},
                {new QL.Date(31, 3.ToQuantLibMonth(), 2025).ToOaDate()},
                {new QL.Date(31, 3.ToQuantLibMonth(), 2026).ToOaDate()},
                {new QL.Date(31, 3.ToQuantLibMonth(), 2027).ToOaDate()},
                {new QL.Date(31, 3.ToQuantLibMonth(), 2028).ToOaDate()}, 
                {new QL.Date(29, 3.ToQuantLibMonth(), 2029).ToOaDate()}, 
                {new QL.Date(29, 3.ToQuantLibMonth(), 2030).ToOaDate()},
                {new QL.Date(31, 3.ToQuantLibMonth(), 2031).ToOaDate()},
                {new QL.Date(31, 3.ToQuantLibMonth(), 2032).ToOaDate()},
                {new QL.Date(31, 3.ToQuantLibMonth(), 2033).ToOaDate()},
                {new QL.Date(30, 3.ToQuantLibMonth(), 2035).ToOaDate()},
                {new QL.Date(31, 3.ToQuantLibMonth(), 2038).ToOaDate()},
                {new QL.Date(31, 3.ToQuantLibMonth(), 2043).ToOaDate()},
                {new QL.Date(31, 3.ToQuantLibMonth(), 2048).ToOaDate()},
                {new QL.Date(31, 3.ToQuantLibMonth(), 2053).ToOaDate()},
            };

        object[,] zarForecastDiscountFactors =
            {
                {1.000000000}, 
                {0.999419242}, 
                {0.999224145}, 
                {0.997834322}, 
                {0.994368225}, 
                {0.987684355}, 
                {0.981464439}, 
                {0.960911793},
                {0.940721043}, 
                {0.920577492}, 
                {0.855851457}, 
                {0.792353567}, 
                {0.729876438}, 
                {0.667805514}, 
                {0.605687930},
                {0.546129323}, 
                {0.490584027}, 
                {0.439981173}, 
                {0.395778224}, 
                {0.317716986}, 
                {0.232076097}, 
                {0.149146864},
                {0.104645248}, 
                {0.046601701},
            };

        string zarForecastCurveHandle = 
            CurveUtils.Create(
                handle: "ZarForecastCurve", 
                curveParameters: zarForecastCurveParameters, 
                datesRange: zarForecastCurveDates, 
                discountFactorsRange: zarForecastDiscountFactors);

        object[,] usdForecastCurveParameters =
        {
            { "Parameter", "Value" },
            { "Calendars", "USD"},
            { "DayCountConvention", "Actual360" },
            { "Interpolation", "Cubic" },
        };
        
        object[,] usdForecastCurveDates =
            {
                {new QL.Date(31, 3.ToQuantLibMonth(), 2023).ToOaDate()}, 
                {new QL.Date(3, 4.ToQuantLibMonth(), 2023).ToOaDate()}, 
                {new QL.Date(4, 4.ToQuantLibMonth(), 2023).ToOaDate()}, 
                {new QL.Date(11, 4.ToQuantLibMonth(), 2023).ToOaDate()}, 
                {new QL.Date(4, 5.ToQuantLibMonth(), 2023).ToOaDate()},
                {new QL.Date(5, 6.ToQuantLibMonth(), 2023).ToOaDate()}, 
                {new QL.Date(5, 7.ToQuantLibMonth(), 2023).ToOaDate()}, 
                {new QL.Date(4, 10.ToQuantLibMonth(), 2023).ToOaDate()}, 
                {new QL.Date(4, 1.ToQuantLibMonth(), 2024).ToOaDate()}, 
                {new QL.Date(4, 4.ToQuantLibMonth(), 2024).ToOaDate()},
                {new QL.Date(4, 4.ToQuantLibMonth(), 2025).ToOaDate()}, 
                {new QL.Date(6, 4.ToQuantLibMonth(), 2026).ToOaDate()}, 
                {new QL.Date(5, 4.ToQuantLibMonth(), 2027).ToOaDate()}, 
                {new QL.Date(4, 4.ToQuantLibMonth(), 2028).ToOaDate()}, 
                {new QL.Date(4, 4.ToQuantLibMonth(), 2029).ToOaDate()},
                {new QL.Date(4, 4.ToQuantLibMonth(), 2030).ToOaDate()}, 
                {new QL.Date(4, 4.ToQuantLibMonth(), 2031).ToOaDate()}, 
                {new QL.Date(5, 4.ToQuantLibMonth(), 2032).ToOaDate()}, 
                {new QL.Date(4, 4.ToQuantLibMonth(), 2033).ToOaDate()}, 
                {new QL.Date(4, 4.ToQuantLibMonth(), 2035).ToOaDate()},
                {new QL.Date(5, 4.ToQuantLibMonth(), 2038).ToOaDate()}, 
                {new QL.Date(6, 4.ToQuantLibMonth(), 2043).ToOaDate()}, 
                {new QL.Date(6, 4.ToQuantLibMonth(), 2048).ToOaDate()}, 
                {new QL.Date(4, 4.ToQuantLibMonth(), 2053).ToOaDate()}, 
                {new QL.Date(4, 4.ToQuantLibMonth(), 2063).ToOaDate()},
                {new QL.Date(4, 4.ToQuantLibMonth(), 2073).ToOaDate()},
            };

        object[,] usdForecastCurveDiscountFactors =
            {
                {1.000000000}, 
                {0.999594748}, 
                {0.999459682}, 
                {0.998485477}, 
                {0.995236637}, 
                {0.990656560}, 
                {0.986389473}, 
                {0.973558139},
                {0.961709557}, 
                {0.951187072}, 
                {0.917197683}, 
                {0.888698879}, 
                {0.862123520}, 
                {0.836076594}, 
                {0.810267688},
                {0.785011688}, 
                {0.760210466}, 
                {0.735227273}, 
                {0.710992469}, 
                {0.663886482}, 
                {0.598545100}, 
                {0.509908045},
                {0.446464292}, 
                {0.395983824}, 
                {0.331760479}, 
                {0.293876050},
            };

        string usdForecastCurveHandle =
            CurveUtils.Create(
                "UsdForecastCurveHandle",
                usdForecastCurveParameters,
                usdForecastCurveDates,
                usdForecastCurveDiscountFactors);
        
        string usdDiscountCurveHandle = 
            CurveUtils.Create(
                handle: "UsdDiscountCurveHandle",
                curveParameters: usdForecastCurveParameters,
                datesRange: usdForecastCurveDates,
                discountFactorsRange: usdForecastCurveDiscountFactors);
        
            // {"FxSpot", 17.7927},
        object[,] usdZarFxBasisCurveParameters =
        {
            {"CurveUtils Parameters", ""},
            {"Parameter", "Value"},
            {"BaseDate", baseDate.ToOaDate()},
            {"BaseCurrencyIndexName", "USD-LIBOR"},
            {"BaseCurrencyIndexTenor", "3m"},
            {"BaseCurrencyForecastCurveHandle", usdForecastCurveHandle},
            {"BaseCurrencyDiscountCurveHandle", usdDiscountCurveHandle}, 
            {"QuoteCurrencyForecastCurveHandle", zarForecastCurveHandle}, 
            {"QuoteCurrencyIndexName", "JIBAR"},
            {"QuoteCurrencyIndexTenor", "3m"},
            {"Interpolation", "Cubic"},
        };
//         fx_swap_spreads: dict[ql.Period, float] = \
//         {
// # ql.Period('1M'): (465.0000 + 484.0000) / 20000,
// # ql.Period('2M'): (953.2000 + 973.2000) / 20000,
//             ql.Period('3M'): 0.14265, # (1411.0000 + 1442.0000) / 20000, #0.1022,
// # ql.Period('4M'): (1888.0000 + 1928.0000) / 20000,
// # ql.Period('5M'): (2404.0000 + 2464.0000) / 20000,
//             ql.Period('6M'): 0.2911,  #0.2468,  # (2871.0000 + 2951.0000) / 20000,
// # ql.Period('7M'): (3417.0000 + 3507.0000) / 20000,
// # ql.Period('8M'): (3871.0700 + 4091.0700) / 20000,
//             ql.Period('9M'): 0.446662, # 0.4108,  # (4409.1200 + 4524.1200) / 20000,
// # ql.Period('10M'): (4916.0000 + 5066.0000) / 20000,
// # ql.Period('11M'): (5377.0000 + 5557.0000) / 20000,
//             ql.Period('1Y'):  0.604391,  # (5946.41 + 6141.41) / 20000,
//         } 

        object[,] fecInstruments =
        {

        };
        
        object[,] crossCurrencySwapInstruments = 
        {
            {"Cross Currency Swaps", "", "", ""},
            {"Tenors", "BasisSpreads", "FixingDays", "Include"},
            {"2y", (1.18 + 1.74)/20000, 2, "TRUE"},
            {"3y", (1.31 + 1.88)/20000, 2, "TRUE"},
            {"4y", (1.37 + 1.94)/20000, 2, "TRUE"},
            {"5y", (1.44 + 2.00)/20000, 2, "TRUE"},
            {"6y", (1.45 + 2.01)/20000, 2, "TRUE"},
            {"7y", (1.32 + 1.88)/20000, 2, "TRUE"},
            {"8y", (1.09 + 1.65)/20000, 2, "TRUE"},
            {"9y", (0.76 + 1.32)/20000, 2, "TRUE"},
            {"10y", (0.37 + 0.93)/20000, 2, "TRUE"},
            {"15y", (-1.37 + -0.80) / 20000, 2, "TRUE"},
            {"20y", (-2.64 + -2.08)/20000, 2, "TRUE"},
        };
// # ql.Period('11Y'): (3.15 + 3.451)/20000.00,
//         ql.Period('12Y'): (-0.42 + 0.14)/20000, #-7.54/10000.00,
//         ql.Period('15Y'): (-1.37 + -0.80) / 20000, # -24.28/10000.00,
//         ql.Period('20Y'): (-2.64 + -2.08)/20000,  #-47.01/10000.00, 

        string usdZarFxBasisCurveHandle = 
            CurveBootstrapper.BootstrapFxBasisCurve(
                handle: "UsdZarFxBasisCurve",
                curveParameters: usdZarFxBasisCurveParameters,
                customRateIndex: null,
                instrumentGroups: crossCurrencySwapInstruments);
        //"1M", "2M", "3M", "4M", "5M", "6M", "9M","1Y", 
        List<string> tenors = new()
        {
            "2Y", "3Y", "4Y", "5Y", "6Y", "7Y", "8Y", "9Y", "10Y", "15Y", "20Y"
        };
        object[] dates = tenors.Select(t => (baseDate + new QL.Period(t)).ToOaDate()).Cast<object>().ToArray(); 
        QL.YieldTermStructure? curve = CurveUtils.GetCurveObject(usdZarFxBasisCurveHandle);
        object[,] actualDiscountFactors = (object[,])CurveUtils.GetDiscountFactors(usdZarFxBasisCurveHandle, dates);
        
        object[,] discountFactorsForExpectedCurve =
        {
            {1.0000000}, 
            {0.9994192}, 
            {0.9992245}, 
            {0.9978405}, 
            {0.9943461}, 
            {0.9871948}, 
            {0.9805455}, 
            {0.9609479}, 
            {0.9416501}, 
            {0.9232669},
            {0.8554395},
            {0.7922287}, 
            {0.7297502}, 
            {0.6676890}, 
            {0.6055767}, 
            {0.5460244}, 
            {0.4904856}, 
            {0.4398893}, 
            {0.3956932}, 
            {0.3176445}, 
            {0.2320204},
            {0.1491120}, 
            {0.1046235}, 
            {0.0465877}
        };
        
        string expectedDiscountCurveHandle = 
            CurveUtils.Create(
                handle: "ExpectedDiscountCurve", 
                curveParameters: zarForecastCurveParameters, 
                datesRange: zarForecastCurveDates, 
                discountFactorsRange: discountFactorsForExpectedCurve);
       
        object[,] expectedDiscountFactors = (object[,])CurveUtils.GetDiscountFactors(expectedDiscountCurveHandle, dates);
          
    }
}
