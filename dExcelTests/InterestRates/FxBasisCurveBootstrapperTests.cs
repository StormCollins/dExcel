using FxCurveBootstrapper = dExcel.FX.CurveBootstrapper;
using dExcel.CommonEnums;
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
        // TODO: Check USDZAR swap spreads
        // TODO: Compare FX forwards in Refinitiv
        QL.Date baseDate = new(31, 3.ToQuantLibMonth(), 2023);

        object[,] zarForecastCurveParameters =
        {
            {"Parameter", "Value"},
            {"Calendars", "ZAR"},
            {"DayCountConvention", "Actual365"},
            {"Interpolation", CurveInterpolationMethods.CubicSpline_On_DiscountFactors.ToString()},
        };
        
        object[,] zarForecastCurveDates =
        {
            {new DateTime(2023, 03, 31).ToOADate()},
            {new DateTime(2023, 04, 03).ToOADate()},
            {new DateTime(2023, 04, 04).ToOADate()},
            {new DateTime(2023, 04, 11).ToOADate()},
            {new DateTime(2023, 04, 28).ToOADate()},
            {new DateTime(2023, 05, 31).ToOADate()},
            {new DateTime(2023, 06, 30).ToOADate()},
            {new DateTime(2023, 09, 29).ToOADate()},
            {new DateTime(2023, 12, 29).ToOADate()},
            {new DateTime(2024, 03, 28).ToOADate()},
            {new DateTime(2025, 03, 31).ToOADate()},
            {new DateTime(2026, 03, 31).ToOADate()},
            {new DateTime(2027, 03, 31).ToOADate()},
            {new DateTime(2028, 03, 31).ToOADate()},
            {new DateTime(2029, 03, 29).ToOADate()},
            {new DateTime(2030, 03, 29).ToOADate()},
            {new DateTime(2031, 03, 31).ToOADate()},
            {new DateTime(2032, 03, 31).ToOADate()},
            {new DateTime(2033, 03, 31).ToOADate()},
            {new DateTime(2035, 03, 30).ToOADate()},
            {new DateTime(2038, 03, 31).ToOADate()},
            {new DateTime(2043, 03, 31).ToOADate()},
            {new DateTime(2048, 03, 31).ToOADate()},
            {new DateTime(2053, 03, 31).ToOADate()},
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
            CurveUtils.CreateFromDiscountFactors(
                handle: "ZarForecastCurve", 
                curveParameters: zarForecastCurveParameters, 
                datesRange: zarForecastCurveDates, 
                discountFactorsRange: zarForecastDiscountFactors);

        object[,] usdForecastCurveParameters =
        {
            { "Parameter", "Value" },
            { "Calendars", "USD"},
            { "DayCountConvention", "Actual360" },
            { "Interpolation", CurveInterpolationMethods.CubicSpline_On_DiscountFactors.ToString() },
        };
        
        object[,] usdForecastCurveDates =
            {
                {new DateTime(2023, 03, 31).ToOADate()},
                {new DateTime(2023, 04, 03).ToOADate()},
                {new DateTime(2023, 04, 04).ToOADate()},
                {new DateTime(2023, 04, 11).ToOADate()}, 
                {new DateTime(2023, 05, 04).ToOADate()},
                {new DateTime(2023, 06, 05).ToOADate()},
                {new DateTime(2023, 07, 05).ToOADate()},
                {new DateTime(2023, 10, 04).ToOADate()},
                {new DateTime(2024, 01, 04).ToOADate()},
                {new DateTime(2024, 04, 04).ToOADate()}, 
                {new DateTime(2025, 04, 04).ToOADate()},
                {new DateTime(2026, 04, 06).ToOADate()},
                {new DateTime(2027, 04, 05).ToOADate()},
                {new DateTime(2028, 04, 04).ToOADate()},
                {new DateTime(2029, 04, 04).ToOADate()},
                {new DateTime(2030, 04, 04).ToOADate()},       
                {new DateTime(2031, 04, 04).ToOADate()},
                {new DateTime(2032, 04, 05).ToOADate()},
                {new DateTime(2033, 04, 04).ToOADate()},
                {new DateTime(2035, 04, 04).ToOADate()},
                {new DateTime(2038, 04, 05).ToOADate()},
                {new DateTime(2043, 04, 06).ToOADate()},
                {new DateTime(2048, 04, 06).ToOADate()},
                {new DateTime(2053, 04, 04).ToOADate()},
                {new DateTime(2063, 04, 04).ToOADate()},
                {new DateTime(2073, 04, 04).ToOADate()},
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
            CurveUtils.CreateFromDiscountFactors(
                "UsdForecastCurveHandle",
                usdForecastCurveParameters,
                usdForecastCurveDates,
                usdForecastCurveDiscountFactors);
        
        string usdDiscountCurveHandle = 
            CurveUtils.CreateFromDiscountFactors(
                handle: "UsdDiscountCurveHandle",
                curveParameters: usdForecastCurveParameters,
                datesRange: usdForecastCurveDates,
                discountFactorsRange: usdForecastCurveDiscountFactors);
        
        object[,] usdZarFxBasisCurveParameters =
        {
            {"CurveUtils Parameters", ""},
            {"Parameter", "Value"},
            {"BaseDate", baseDate.ToOaDate()},
            {"SpotFx", 17.7927},
            {"BaseCurrencyIndexName", RateIndices.USD_LIBOR.ToString()},
            {"BaseCurrencyIndexTenor", "3m"},
            {"BaseCurrencyForecastCurveHandle", usdForecastCurveHandle},
            {"BaseCurrencyDiscountCurveHandle", usdDiscountCurveHandle}, 
            {"QuoteCurrencyForecastCurveHandle", zarForecastCurveHandle}, 
            {"QuoteCurrencyIndexName", "JIBAR"},
            {"QuoteCurrencyIndexTenor", "3m"},
            {"Interpolation", CurveInterpolationMethods.CubicSpline_On_DiscountFactors.ToString()},
        };

        object[,] fecInstruments =
        {
            {"FECs", "", "", ""},
            {"Tenors", "ForwardPoints", "FixingDays", "Include"},
            {"3M", 0.142650, 2, "TRUE"},
            {"6M", 0.291100, 2, "TRUE"},
            {"9M", 0.446662, 2, "TRUE"},
            {"1Y", 0.604391, 2, "TRUE"},
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
            {"12y", (-0.42 + 0.14)/20000, 2, "TRUE"},
            {"15y", (-1.37 + -0.80) / 20000, 2, "TRUE"},
            {"20y", (-2.64 + -2.08)/20000, 2, "TRUE"},
        };
        
        string usdZarFxBasisCurveHandle = 
            FxCurveBootstrapper.BootstrapFxBasisCurve(
                handle: "UsdZarFxBasisCurve",
                curveParameters: usdZarFxBasisCurveParameters,
                customBaseCurrencyIndex: null,
                customQuoteCurrencyIndex: null,
                instrumentGroups: new object[] {fecInstruments, crossCurrencySwapInstruments});
        
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
            CurveUtils.CreateFromDiscountFactors(
                handle: "ExpectedDiscountCurve", 
                curveParameters: zarForecastCurveParameters, 
                datesRange: zarForecastCurveDates, 
                discountFactorsRange: discountFactorsForExpectedCurve);
       
        object[,] expectedDiscountFactors = (object[,])CurveUtils.GetDiscountFactors(expectedDiscountCurveHandle, dates);
    }

    [Test]
    public void RefinitivFxForwardsTest()
    {
        object[,] usdDiscountCurveParameters =
        {
            { "Parameter", "Value" },
            { "Calendars", "USD"},
            { "DayCountConvention", "Actual360" },
            { "Interpolation", CurveInterpolationMethods.CubicSpline_On_DiscountFactors.ToString() },
        };
        
        object[,] usdDiscountCurveDates =
            {
                {new DateTime(2023, 03, 31).ToOADate()},
                {new DateTime(2023, 04, 03).ToOADate()},
                {new DateTime(2023, 04, 04).ToOADate()},
                {new DateTime(2023, 04, 11).ToOADate()}, 
                {new DateTime(2023, 05, 04).ToOADate()},
                {new DateTime(2023, 06, 05).ToOADate()},
                {new DateTime(2023, 07, 05).ToOADate()},
                {new DateTime(2023, 10, 04).ToOADate()},
                {new DateTime(2024, 01, 04).ToOADate()},
                {new DateTime(2024, 04, 04).ToOADate()}, 
                {new DateTime(2025, 04, 04).ToOADate()},
                {new DateTime(2026, 04, 06).ToOADate()},
                {new DateTime(2027, 04, 05).ToOADate()},
                {new DateTime(2028, 04, 04).ToOADate()},
                {new DateTime(2029, 04, 04).ToOADate()},
                {new DateTime(2030, 04, 04).ToOADate()},       
                {new DateTime(2031, 04, 04).ToOADate()},
                {new DateTime(2032, 04, 05).ToOADate()},
                {new DateTime(2033, 04, 04).ToOADate()},
                {new DateTime(2035, 04, 04).ToOADate()},
                {new DateTime(2038, 04, 05).ToOADate()},
                {new DateTime(2043, 04, 06).ToOADate()},
                {new DateTime(2048, 04, 06).ToOADate()},
                {new DateTime(2053, 04, 04).ToOADate()},
                {new DateTime(2063, 04, 04).ToOADate()},
                {new DateTime(2073, 04, 04).ToOADate()},
            };

        object[,] usdDiscountCurveDiscountFactors =
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

        string usdDiscountCurveHandle = 
            CurveUtils.CreateFromDiscountFactors(
                handle: "UsdDiscountCurveHandle",
                curveParameters: usdDiscountCurveParameters,
                datesRange: usdDiscountCurveDates,
                discountFactorsRange: usdDiscountCurveDiscountFactors);
        
        object[,] usdZarFxBasisCurveDiscountFactors =
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
        
        object[,] usdZarFxBasisCurveDates =
        {
            {new DateTime(2023, 03, 31).ToOADate()},
            {new DateTime(2023, 04, 03).ToOADate()},
            {new DateTime(2023, 04, 04).ToOADate()},
            {new DateTime(2023, 04, 11).ToOADate()},
            {new DateTime(2023, 04, 28).ToOADate()},
            {new DateTime(2023, 05, 31).ToOADate()},
            {new DateTime(2023, 06, 30).ToOADate()},
            {new DateTime(2023, 09, 29).ToOADate()},
            {new DateTime(2023, 12, 29).ToOADate()},
            {new DateTime(2024, 03, 28).ToOADate()},
            {new DateTime(2025, 03, 31).ToOADate()},
            {new DateTime(2026, 03, 31).ToOADate()},
            {new DateTime(2027, 03, 31).ToOADate()},
            {new DateTime(2028, 03, 31).ToOADate()},
            {new DateTime(2029, 03, 29).ToOADate()},
            {new DateTime(2030, 03, 29).ToOADate()},
            {new DateTime(2031, 03, 31).ToOADate()},
            {new DateTime(2032, 03, 31).ToOADate()},
            {new DateTime(2033, 03, 31).ToOADate()},
            {new DateTime(2035, 03, 30).ToOADate()},
            {new DateTime(2038, 03, 31).ToOADate()},
            {new DateTime(2043, 03, 31).ToOADate()},
            {new DateTime(2048, 03, 31).ToOADate()},
            {new DateTime(2053, 03, 31).ToOADate()},
        };
        
        object[,] usdZarFxBasisCurveParameters = 
        {
            {"Parameter", "Value"},
            {"Calendars", "ZAR"},
            {"DayCountConvention", "Actual365"},
            {"Interpolation", CurveInterpolationMethods.CubicSpline_On_DiscountFactors.ToString()},
        };
        
        string usdZarFxBasisDiscountCurveHandle = 
            CurveUtils.CreateFromDiscountFactors(
                handle: "ExpectedDiscountCurve", 
                curveParameters: usdZarFxBasisCurveParameters, 
                datesRange: usdZarFxBasisCurveDates, 
                discountFactorsRange: usdZarFxBasisCurveDiscountFactors);

        Dictionary<string, double> usdZarFxRates = 
            new()
            {
                ["0D"] = 17.7927,
                ["1D"] = 17.7942,
                ["1W"] = 17.8034,
                ["2W"] = 17.8143,
                ["3W"] = 17.8254,
                ["1M"] = 17.8402,
                ["2M"] = 17.8890,
                ["3M"] = 17.9354,
                ["4M"] = 17.9835,
                ["5M"] = 18.0361,
                ["6M"] = 18.0838,
                ["7M"] = 18.1389,
                ["8M"] = 18.1908,
                ["9M"] = 18.2394,
                ["10M"] = 18.2918,
                ["11M"] = 18.3394,
                ["1Y"] = 18.4546,
                ["2Y"] = 19.1752,
                ["3Y"] = 20.1475,
                ["4Y"] = 21.2794,
                ["5Y"] = 22.6297,
                ["10Y"] = 31.9600,
                ["15Y"] = 42.6219,
                ["20Y"] = 50.8479,
            };
        
        
        QL.Date baseDate = new(31, 3.ToQuantLibMonth(), 2023);
        object[] dates = {baseDate + new QL.Period("1Y")};
        dates = dates.Select(x => ((QL.Date)x).ToOaDate()).Cast<object>().ToArray();
            
        double usdDiscountFactor = 
                    (double)((object[,])CurveUtils.GetDiscountFactors(usdDiscountCurveHandle, dates))[0, 0];
        
        double usdZarFxBasisDiscountFactor = 
            (double)((object[,])CurveUtils.GetDiscountFactors(usdZarFxBasisDiscountCurveHandle, dates))[0, 0];
        
        double fxForward =  usdZarFxRates["0D"] * usdDiscountFactor / usdZarFxBasisDiscountFactor; 

    }

    [Test]
    public void usdZarFxBasisCurvedExcelVsOldValsMethodTest()
    {

        DateTime baseDate = new(2021, 03, 31);
        
        object[,] usdOisCurveParameters =
        {
            {"Parameter", "Value"},
            {"Calendars", "USD"},
            {"DayCountConvention", "Actual360"},
            {"Interpolation", CurveInterpolationMethods.Exponential_On_DiscountFactors.ToString()},
        };
         
         object[,] usdOisCurveDates =
             {
                 {new DateTime(2021, 03, 31).ToOADate()},
                 {new DateTime(2021, 04, 05).ToOADate()},
                 {new DateTime(2021, 04, 06).ToOADate()},
                 {new DateTime(2021, 05, 05).ToOADate()},
                 {new DateTime(2021, 06, 07).ToOADate()},
                 {new DateTime(2021, 07, 06).ToOADate()},
                 {new DateTime(2021, 08, 05).ToOADate()},
                 {new DateTime(2021, 09, 07).ToOADate()},
                 {new DateTime(2021, 10, 05).ToOADate()},
                 {new DateTime(2022, 01, 05).ToOADate()},
                 {new DateTime(2022, 04, 05).ToOADate()},
                 {new DateTime(2022, 10, 05).ToOADate()},
                 {new DateTime(2023, 04, 05).ToOADate()},
                 {new DateTime(2024, 04, 05).ToOADate()},
                 {new DateTime(2025, 04, 07).ToOADate()},
                 {new DateTime(2026, 04, 06).ToOADate()},
                 {new DateTime(2028, 04, 05).ToOADate()},
                 {new DateTime(2031, 04, 07).ToOADate()},
                 {new DateTime(2033, 04, 05).ToOADate()},
                 {new DateTime(2036, 04, 07).ToOADate()},
                 {new DateTime(2041, 04, 05).ToOADate()},
                 {new DateTime(2046, 04, 05).ToOADate()},
                 {new DateTime(2051, 04, 05).ToOADate()},
             };
         
        object[,] usdOisCurveDiscountFactors =
            {
                {1.000000000000000},
                {0.999991666708333},
                {0.999990000058333},
                {0.999940697710193},
                {0.999879125186333},
                {0.999819503967011},
                {0.999753115322318},
                {0.999680100768426},
                {0.999618921121899},
                {0.999409234979339},
                {0.999178448712571},
                {0.998471498030183},
                {0.997042585805172},
                {0.989955122288159},
                {0.976382944031500},
                {0.958598404601822},
                {0.916847221922653},
                {0.853134704048841},
                {0.811853567594985},
                {0.754613299955985},
                {0.672596673719967},
                {0.604138880680194},
                {0.545474652384628},
            };
        
        string usdOisCurveHandle = 
            CurveUtils.CreateFromDiscountFactors(
                handle: "UsdOisCurveHandle",
                curveParameters: usdOisCurveParameters,
                datesRange: usdOisCurveDates,
                discountFactorsRange: usdOisCurveDiscountFactors);

        object[,] usdSwapCurveParameters =
        {
            { "Parameter", "Value" },
            { "Calendars", "USD"},
            { "DayCountConvention", "Actual360" },
            { "Interpolation", CurveInterpolationMethods.CubicSpline_On_DiscountFactors.ToString() },
        };
        
        object[,] usdSwapCurveDates =
        {
            {new DateTime(2021, 03, 31).ToOADate()},
            {new DateTime(2021, 04, 01).ToOADate()},
            {new DateTime(2021, 04, 12).ToOADate()},
            {new DateTime(2021, 05, 05).ToOADate()},
            {new DateTime(2021, 06, 07).ToOADate()},
            {new DateTime(2021, 07, 06).ToOADate()},
            {new DateTime(2021, 07, 30).ToOADate()},
            {new DateTime(2021, 09, 01).ToOADate()},
            {new DateTime(2021, 09, 30).ToOADate()},
            {new DateTime(2021, 11, 02).ToOADate()},
            {new DateTime(2021, 11, 30).ToOADate()},
            {new DateTime(2021, 12, 30).ToOADate()},
            {new DateTime(2022, 02, 01).ToOADate()},
            {new DateTime(2022, 02, 28).ToOADate()},
            {new DateTime(2022, 03, 31).ToOADate()},
            {new DateTime(2022, 06, 30).ToOADate()},
            {new DateTime(2022, 09, 30).ToOADate()},
            {new DateTime(2022, 12, 30).ToOADate()},
            {new DateTime(2023, 04, 03).ToOADate()},
            {new DateTime(2024, 04, 05).ToOADate()},
            {new DateTime(2025, 04, 07).ToOADate()},
            {new DateTime(2026, 04, 06).ToOADate()},
            {new DateTime(2027, 04, 05).ToOADate()},
            {new DateTime(2028, 04, 05).ToOADate()},
            {new DateTime(2029, 04, 05).ToOADate()},
            {new DateTime(2030, 04, 05).ToOADate()},
            {new DateTime(2031, 04, 07).ToOADate()},
            {new DateTime(2033, 04, 05).ToOADate()},
            {new DateTime(2036, 04, 07).ToOADate()},
            {new DateTime(2041, 04, 05).ToOADate()},
            {new DateTime(2046, 04, 05).ToOADate()},
            {new DateTime(2051, 04, 05).ToOADate()},
            {new DateTime(2061, 04, 05).ToOADate()},
        };

        object[,] usdSwapCurveDiscountFactors =
        {
            {1.000000000000000},
            {0.999997790963784},
            {0.999971233704236},
            {0.999893448340767},
            {0.999751107442379},
            {0.999484040324416},
            {0.999463824633386},
            {0.999338700785530},
            {0.999100101755088},
            {0.998999204520188},
            {0.998869329154220},
            {0.998574796695599},
            {0.998405844300001},
            {0.998234288053710},
            {0.997917612531781},
            {0.997330799483844},
            {0.996570697265890},
            {0.995610083558693},
            {0.994323858297453},
            {0.984327271846949},
            {0.968442772651898},
            {0.948053193400618},
            {0.925525740023929},
            {0.902521304652302},
            {0.879821170547571},
            {0.857059855744345},
            {0.834258179767811},
            {0.790405775065646},
            {0.729867677573608},
            {0.642791894309980},
            {0.572132772719208},
            {0.510944937056050},
            {0.425373409630619},
        };

        string usdSwapCurveHandle = 
            CurveUtils.CreateFromDiscountFactors(
                handle: "UsdSwapCurveHandle",
                curveParameters: usdSwapCurveParameters,
                datesRange: usdSwapCurveDates,
                discountFactorsRange: usdSwapCurveDiscountFactors);
        
        object[,] zarSwapCurveParameters =
        {
            { "Parameter", "Value" },
            { "Calendars", "ZAR"},
            { "DayCountConvention", "Actual365" },
            { "Interpolation", CurveInterpolationMethods.CubicSpline_On_DiscountFactors.ToString() },
        };
        
        object[,] zarSwapCurveDates =
        {
            {new DateTime(2021, 03, 31).ToOADate()},
            {new DateTime(2021, 04, 01).ToOADate()},
            {new DateTime(2021, 04, 30).ToOADate()},
            {new DateTime(2021, 06, 30).ToOADate()},
            {new DateTime(2021, 09, 30).ToOADate()},
            {new DateTime(2021, 12, 30).ToOADate()},
            {new DateTime(2022, 03, 31).ToOADate()},
            {new DateTime(2022, 06, 30).ToOADate()},
            {new DateTime(2022, 09, 30).ToOADate()},
            {new DateTime(2022, 12, 30).ToOADate()},
            {new DateTime(2023, 04, 03).ToOADate()},
            {new DateTime(2024, 03, 28).ToOADate()},
            {new DateTime(2025, 03, 31).ToOADate()},
            {new DateTime(2026, 03, 31).ToOADate()},
            {new DateTime(2027, 03, 31).ToOADate()},
            {new DateTime(2028, 03, 31).ToOADate()},
            {new DateTime(2029, 03, 29).ToOADate()},
            {new DateTime(2030, 03, 29).ToOADate()},
            {new DateTime(2031, 03, 31).ToOADate()},
            {new DateTime(2033, 03, 31).ToOADate()},
            {new DateTime(2036, 03, 31).ToOADate()},
            {new DateTime(2041, 03, 29).ToOADate()},
            {new DateTime(2046, 03, 30).ToOADate()},
            {new DateTime(2051, 03, 31).ToOADate()},
        };

        object[,] zarSwapCurveDiscountFactors =
        {
            {1.000000000000000},
            {0.999909323303551},
            {0.997125001782141},
            {0.990920857293341},
            {0.981544285721100},
            {0.971804409303220},
            {0.961567204903394},
            {0.950570793608732},
            {0.939114115961109},
            {0.927254656104601},
            {0.914695369530117},
            {0.862343748830009},
            {0.802112116095088},
            {0.738112493430834},
            {0.674055878801053},
            {0.610952848474239},
            {0.551394511492711},
            {0.495515125227156},
            {0.443568513868495},
            {0.351549327385578},
            {0.248554400584266},
            {0.141976366361454},
            {0.082442215018176},
            {0.053272385379469},
        };

        string zarSwapCurveHandle = 
            CurveUtils.CreateFromDiscountFactors(
                handle: "ZarSwapCurveHandle",
                curveParameters: zarSwapCurveParameters,
                datesRange: zarSwapCurveDates,
                discountFactorsRange: zarSwapCurveDiscountFactors);
       
        object[,] usdZarFxBasisCurveParameters =
        {
            {"CurveUtils Parameters", ""},
            {"Parameter", "Value"},
            {"BaseDate", baseDate.ToOADate()},
            {"SpotFx", 14.7768},
            {"BaseCurrencyIndexName", RateIndices.USD_LIBOR.ToString()},
            {"BaseCurrencyIndexTenor", "3m"},
            {"BaseCurrencyForecastCurveHandle", usdSwapCurveHandle},
            {"BaseCurrencyDiscountCurveHandle", usdOisCurveHandle}, 
            {"QuoteCurrencyForecastCurveHandle", zarSwapCurveHandle}, 
            {"QuoteCurrencyIndexName", "JIBAR"},
            {"QuoteCurrencyIndexTenor", "3m"},
            {"Interpolation", CurveInterpolationMethods.Exponential_On_DiscountFactors.ToString()},
        };
        
        object[,] crossCurrencySwapInstruments = 
        {
            {"Cross Currency Swaps", "", "", ""},
            {"Tenors", "BasisSpreads", "FixingDays", "Include"},
            {"2y", 69/10_000, 2, "TRUE"},
            {"3y", 57/10_000, 2, "TRUE"},
            {"4y", 46/10_000, 2, "TRUE"},
            {"5y", 38/10_000, 2, "TRUE"},
            {"6y", 30/10_000, 2, "TRUE"},
            {"7y", 23/10_000, 2, "TRUE"},
            {"8y", 17/10_000, 2, "TRUE"},
            {"9y", 11/10_000, 2, "TRUE"},
            {"10y", 6/10_000, 2, "TRUE"},
            {"12y", 0/10_000, 2, "TRUE"},
            {"15y", -8/10_000, 2, "TRUE"},
            {"20y", -23.5/10_000, 2, "TRUE"},
        };

        string usdZarFxBasisCurveHandle = 
            dExcel.FX.CurveBootstrapper.BootstrapFxBasisCurve(
                handle: "OldValsMethod",
                curveParameters: usdZarFxBasisCurveParameters,
                customBaseCurrencyIndex: null,
                customQuoteCurrencyIndex: null,
                instrumentGroups: crossCurrencySwapInstruments);
        
        object[] usdZarFxBasisCurveDates =
        {
            new DateTime(2021, 03, 31).ToOADate(),
            // {new DateTime(2021, 04, 01)},
            new DateTime(2021, 04, 06).ToOADate(),
            new DateTime(2021, 04, 07).ToOADate(),
            new DateTime(2021, 04, 13).ToOADate(),
            new DateTime(2021, 04, 20).ToOADate(),
            new DateTime(2021, 04, 28).ToOADate(),
            new DateTime(2021, 05, 06).ToOADate(),
            new DateTime(2021, 06, 07).ToOADate(),
            new DateTime(2021, 07, 06).ToOADate(),
            new DateTime(2021, 08, 06).ToOADate(),
            new DateTime(2021, 09, 07).ToOADate(),
            new DateTime(2021, 10, 06).ToOADate(),
            new DateTime(2022, 01, 06).ToOADate(),
            new DateTime(2022, 04, 06).ToOADate(),
            new DateTime(2023, 04, 03).ToOADate(),
            new DateTime(2024, 04, 02).ToOADate(),
            new DateTime(2025, 04, 01).ToOADate(),
            new DateTime(2026, 04, 01).ToOADate(),
            new DateTime(2027, 04, 01).ToOADate(),
            new DateTime(2028, 04, 03).ToOADate(),
            new DateTime(2029, 04, 03).ToOADate(),
            new DateTime(2030, 04, 01).ToOADate(),
            new DateTime(2031, 04, 01).ToOADate(),
            new DateTime(2033, 04, 01).ToOADate(),
            new DateTime(2036, 04, 01).ToOADate(),
            new DateTime(2041, 04, 01).ToOADate(),
        };

        object[,] expectedUsdZarFxBasisCurveDiscountFactors =
        {
            {1.000000000000000},
            {0.999020338954623},
            {0.998869212808448},
            {0.997945433114428},
            {0.997032239707795},
            {0.995924891169625},
            {0.994843217963034},
            {0.990624509442702},
            {0.986645413250824},
            {0.982723450907178},
            {0.978613573383738},
            {0.974851567153882},
            {0.963200447599801},
            {0.951915120620261},
            {0.904204013382439},
            {0.852678872506256},
            {0.795307483450813},
            {0.734429961861639},
            {0.673568117531598},
            {0.614565233927124},
            {0.557807917552175},
            {0.506432808631880},
            {0.458797458113214},
            {0.370655416421902},
            {0.272278427149989},
            {0.174006444590892},
        };

        object[,] actualUsdZarFxBasisCurveDiscountFactors  = 
            (object[,])CurveUtils.GetDiscountFactors(usdZarFxBasisCurveHandle, usdZarFxBasisCurveDates);

        for (int i = 0; i < expectedUsdZarFxBasisCurveDiscountFactors.Length; i++)
        {
            Assert.AreEqual(
                expected: (double)expectedUsdZarFxBasisCurveDiscountFactors[i, 0],
                actual: (double)actualUsdZarFxBasisCurveDiscountFactors[i, 0],
                delta: 0.05);
        } 
    }
}
