using dExcel.CommonEnums;
using dExcel.InterestRates;
using NUnit.Framework;

namespace dExcelTests.FX;

[TestFixture]
public class CurveBootstrapperTests
{
   /// <summary>
   /// Here we compare the QuantLib FX basis curve bootstrapping for USDZAR (using USD-OIS discounting and USD-Swap
   /// forecasting) to that of the old Vals bootstrapping template.
   /// </summary>
   [Test]
   public void BootstrapFxCurveTest()
   {
      object[,] usdOisCurveParameters = 
      {
         {"Parameter", "Value"},
         {"Calendars", "USD"},
         {"DayCountConvention", "Actual360"},
         {"Interpolation", CurveInterpolationMethods.Exponential_On_DiscountFactors.ToString()},
      };
		
      object[,] usdOisDates =
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

      object[,] usdOisDiscountFactors =
      {
         {1.000000},
         {0.999992},
         {0.999990},
         {0.999941},
         {0.999879},
         {0.999820},
         {0.999753},
         {0.999680},
         {0.999619},
         {0.999409},
         {0.999178},
         {0.998471},
         {0.997043},
         {0.989955},
         {0.976383},
         {0.958598},
         {0.916847},
         {0.853135},
         {0.811854},
         {0.754613},
         {0.672597},
         {0.604139},
         {0.545475},
      };

      string usdOisCurveHandle = 
         CurveUtils.CreateFromDiscountFactors("UsdOisCurve", usdOisCurveParameters, usdOisDates, usdOisDiscountFactors);

      object[,] usdSwapCurveParameters = 
      {
         {"Parameter", "Value"},
         {"Calendars", "USD"},
         {"DayCountConvention", "Actual360"},
         {"Interpolation", CurveInterpolationMethods.Exponential_On_DiscountFactors.ToString()},
      };
		
      object[,] usdSwapDates =
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
      
      object[,] usdSwapDiscountFactors = 
      {
         {1.000000},
         {0.999998},
         {0.999971},
         {0.999893},
         {0.999751},
         {0.999484},
         {0.999464},
         {0.999339},
         {0.999100},
         {0.998999},
         {0.998869},
         {0.998575},
         {0.998406},
         {0.998234},
         {0.997918},
         {0.997331},
         {0.996571},
         {0.995610},
         {0.994324},
         {0.984327},
         {0.968443},
         {0.948053},
         {0.925526},
         {0.902521},
         {0.879821},
         {0.857060},
         {0.834258},
         {0.790406},
         {0.729868},
         {0.642792},
         {0.572133},
         {0.510945},
         {0.425373},
      };
      
      string usdSwapCurveHandle = 
         CurveUtils.CreateFromDiscountFactors(
            handle: "UsdSwapCurve", 
            curveParameters: usdSwapCurveParameters, 
            datesRange: usdSwapDates, 
            discountFactorsRange: usdSwapDiscountFactors);
      
      object[,] zarSwapCurveParameters = 
      {
         {"Parameter", "Value"},
         {"Calendars", "ZAR"},
         {"DayCountConvention", "Actual365"},
         {"Interpolation", CurveInterpolationMethods.Exponential_On_DiscountFactors.ToString()},
      };
		
      object[,] zarSwapDates =
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
      
      object[,] zarSwapDiscountFactors = 
      {
         {1.000000},
         {0.999909},
         {0.997125},
         {0.990921},
         {0.981544},
         {0.971804},
         {0.961567},
         {0.950571},
         {0.939114},
         {0.927255},
         {0.914695},
         {0.862344},
         {0.802112},
         {0.738112},
         {0.674056},
         {0.610953},
         {0.551395},
         {0.495515},
         {0.443569},
         {0.351549},
         {0.248554},
         {0.141976},
         {0.082442},
         {0.053272},
      };
      
      string zarSwapCurveHandle = 
         CurveUtils.CreateFromDiscountFactors(
            handle: "ZarSwapCurve", 
            curveParameters: zarSwapCurveParameters, 
            datesRange: zarSwapDates, 
            discountFactorsRange: zarSwapDiscountFactors);
      
      
      object[,] usdZarFxBasisCurveParameters = 
      {
         {"Parameter", "Value"},
         {"BaseDate", new DateTime(2021, 03, 31).ToOADate()},
         {"BaseCurrencyIndexName", RateIndices.USD_LIBOR.ToString()},
         {"BaseCurrencyIndexTenor", "3M"},
         {"BaseCurrencyDiscountCurve", usdOisCurveHandle},
         {"BaseCurrencyForecastCurve", usdSwapCurveHandle},
         {"QuoteCurrencyIndexName", RateIndices.JIBAR.ToString()},
         {"QuoteCurrencyIndexTenor", "3M"}, 
         {"QuoteCurrencyForecastCurve", zarSwapCurveHandle},
         {"SpotFx", 14.7768},
         {"Interpolation", CurveInterpolationMethods.Exponential_On_DiscountFactors.ToString()},
      };

      object[,] crossCurrencySwaps =
      {
         {"Cross Currency Swaps", "", "", ""},
         {"Tenors", "BasisSpreads", "FixingDays", "Include"},
         {"2Y", 0.00690, 2, "TRUE"},
         {"3Y", 0.00570, 2, "TRUE"},
         {"4Y", 0.00460, 2, "TRUE"},
         {"5Y", 0.00380, 2, "TRUE"},
         {"6Y", 0.00300, 2, "TRUE"},
         {"7Y", 0.00230, 2, "TRUE"},
         {"8Y", 0.00170, 2, "TRUE"},
         {"9Y", 0.00110, 2, "TRUE"},
         {"10Y", 0.00060, 2, "TRUE"},
         {"12Y", 0.00000, 2, "TRUE"},
         {"15Y", -0.00080, 2, "TRUE"},
         {"20Y", -0.00235, 2, "TRUE"},
      };

      string usdZarFxBasisCurveHandle = 
         dExcel.FX.CurveBootstrapper.BootstrapFxBasisAdjustedCurve(
            handle: "UsdZarFxBasisCurve", 
            curveParameters: usdZarFxBasisCurveParameters, 
            customBaseCurrencyIndex: null, 
            customQuoteCurrencyIndex: null, 
            instrumentGroups: crossCurrencySwaps);

      object[] usdZarFxBasisCurveDates =
      {
         new DateTime(2021, 03, 31).ToOADate(),
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
      
      object[,] actualUsdZarFxBasisCurveDiscountFactors = 
         (object[,])CurveUtils.GetDiscountFactors(usdZarFxBasisCurveHandle, usdZarFxBasisCurveDates);

      object[,] expectedUsdZarFxBasisCurveDiscountFactors =
      {
         {1.000000000},
         {0.999020339},
         {0.998869213},
         {0.997945433},
         {0.997032240},
         {0.995924891},
         {0.994843218},
         {0.990624509},
         {0.986645413},
         {0.982723451},
         {0.978613573},
         {0.974851567},
         {0.963200448},
         {0.951915121},
         {0.904204013},
         {0.852678873},
         {0.795307483},
         {0.734429962},
         {0.673568118},
         {0.614565234},
         {0.557807918},
         {0.506432809},
         {0.458797458},
         {0.370655416},
         {0.272278427},
         {0.174006445},
      };

      for (int i = 0; i < expectedUsdZarFxBasisCurveDiscountFactors.GetLength(0); i++)
      {
         Assert.AreEqual(
            expected: (double)expectedUsdZarFxBasisCurveDiscountFactors[i, 0], 
            actual: (double)actualUsdZarFxBasisCurveDiscountFactors[i, 0],
            5e-3); 
      } 
   }
}
