using dExcel.Dates;
using dExcel.FX;
using dExcel.InterestRates;
using NUnit.Framework;
using QL = QuantLib;

namespace dExcelTests.InterestRates;

using CurveBootstrapper = dExcel.FX.CurveBootstrapper;

[TestFixture]
public class FxBasisCurveBootstrapperTests
{
    [Test]
    public void BootstrapFxBasisCurveTest()
    {
        QL.Date baseDate = new(31, 3.ToQuantLibMonth(), 2023);

        object[,] zarForecastCurveParameters =
        {
            {"Parameter", "Value"},
            {"Calendars", "ZAR"},
            {"DayCountConvention", "Actual365"},
            {"Interpolation", "Cubic"},
        };

        object[,] zarForecastCurveDates =
        {
            {new DateTime(2023, 03, 31).ToOADate()},
            {new DateTime(2023, 04, 03).ToOADate()},
            {new DateTime(2023, 04, 04).ToOADate()},
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
            {new DateTime(2035, 03, 31).ToOADate()},
            {new DateTime(2038, 03, 31).ToOADate()},
            {new DateTime(2043, 03, 31).ToOADate()},
            {new DateTime(2048, 03, 31).ToOADate()},
            {new DateTime(2050, 03, 31).ToOADate()},
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
        
        object[,] usdZarFxBasisCurveParameters =
        {
            {"CurveUtils Parameters", ""},
            {"Parameter", "Value"},
            {"BaseDate", baseDate.ToOaDate()},
            {"SpotFx", 17.7927},
            {"BaseCurrencyIndexName", "USD-LIBOR"},
            {"BaseCurrencyIndexTenor", "3m"},
            {"BaseCurrencyForecastCurveHandle", usdForecastCurveHandle},
            {"BaseCurrencyDiscountCurveHandle", usdDiscountCurveHandle}, 
            {"QuoteCurrencyForecastCurveHandle", zarForecastCurveHandle}, 
            {"QuoteCurrencyIndexName", "JIBAR"},
            {"QuoteCurrencyIndexTenor", "3m"},
            {"Interpolation", "Cubic"},
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
            CurveBootstrapper.BootstrapFxBasisCurve(
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
            CurveUtils.Create(
                handle: "ExpectedDiscountCurve", 
                curveParameters: zarForecastCurveParameters, 
                datesRange: zarForecastCurveDates, 
                discountFactorsRange: discountFactorsForExpectedCurve);
       
        object[,] expectedDiscountFactors = (object[,])CurveUtils.GetDiscountFactors(expectedDiscountCurveHandle, dates);
          
    }
}
