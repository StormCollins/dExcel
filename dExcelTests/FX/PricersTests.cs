using dExcel.FX;
using dExcel.Utilities;
using NUnit.Framework;
using mnd = MathNet.Numerics.Distributions;

namespace dExcelTests.FX;

[TestFixture]
public class PricersTests
{
    [Test]
    public void GarmanKohlhagenSpotOptionPricerTest()
    {
         // See Example 17.2 of John Hull - Options, Futures, and Other Derivatives, 9th Edition. 
         // Call option price is 4.3c (approx).
         const double fxSpotPrice = 1.6;
         const double strike = 1.6;
         const double domesticRiskFreeRate = 0.08;
         const double foreignRiskFreeRate = 0.11;
         const double vol = 0.141;
         const double optionMaturity = 4.0 / 12.0;
         double actualCallOptionPrice =
             (double)Pricers.GarmanKohlhagenSpotOptionPricer(
                 spotPrice: fxSpotPrice, 
                 strike: strike, 
                 domesticRiskFreeRate: domesticRiskFreeRate, 
                 foreignRiskFreeRate: foreignRiskFreeRate, 
                 vol: vol, 
                 optionMaturity: optionMaturity, 
                 optionType: "C", 
                 direction: "L", 
                 outputType: "PRICE");
        
        double d1 = (Math.Log(fxSpotPrice / strike) + (domesticRiskFreeRate - foreignRiskFreeRate + Math.Pow(vol, 2)/2) * optionMaturity) / (vol * Math.Sqrt(optionMaturity));
        double d2 = d1 - vol * Math.Sqrt(optionMaturity);
        double domesticDiscountFactor = Math.Exp(-1 * domesticRiskFreeRate * optionMaturity);
        double foreignDiscountFactor = Math.Exp(-1 * foreignRiskFreeRate * optionMaturity);
        double expectedCallOptionPrice = 
                fxSpotPrice * foreignDiscountFactor * mnd.Normal.CDF(0, 1, d1) -
                strike * domesticDiscountFactor * mnd.Normal.CDF(0, 1, d2);
        
        Assert.AreEqual(expectedCallOptionPrice, actualCallOptionPrice);
    }

    [Test]
    public void GarmanKohlhagenSpotOptionPricerForCallOptionVerboseOutputTest()
    {
         // See Example 17.2 of John Hull - Options, Futures, and Other Derivatives, 9th Edition. 
         // Call option price is 4.3c (approx).
         const double fxSpotPrice = 1.6;
         const double strike = 1.6;
         const double domesticRiskFreeRate = 0.08;
         const double foreignRiskFreeRate = 0.11;
         const double vol = 0.141;
         const double optionMaturity = 4.0 / 12.0;
         
         object[,] actual =
             (object[,])Pricers.GarmanKohlhagenSpotOptionPricer(
                 spotPrice: fxSpotPrice, 
                 strike: strike, 
                 domesticRiskFreeRate: domesticRiskFreeRate, 
                 foreignRiskFreeRate: foreignRiskFreeRate, 
                 vol: vol, 
                 optionMaturity: optionMaturity, 
                 optionType: "C", 
                 direction: "L", 
                 outputType: "VERBOSE");

        double d1 = (Math.Log(fxSpotPrice / strike) + (domesticRiskFreeRate - foreignRiskFreeRate + Math.Pow(vol, 2)/2) * optionMaturity) / (vol * Math.Sqrt(optionMaturity));
        double d2 = d1 - vol * Math.Sqrt(optionMaturity);
        double domesticDiscountFactor = Math.Exp(-1 * domesticRiskFreeRate * optionMaturity);
        double foreignDiscountFactor = Math.Exp(-1 * foreignRiskFreeRate * optionMaturity);
        double expectedCallOptionPrice = 
                fxSpotPrice * foreignDiscountFactor * mnd.Normal.CDF(0, 1, d1) -
                strike * domesticDiscountFactor * mnd.Normal.CDF(0, 1, d2);
        
        object[,] expected = 
        {
            {"Price", expectedCallOptionPrice},
            {"d1", d1},
            {"d2", d2},
            {"Domestic Discount Factor", domesticDiscountFactor},
            {"Foreign Discount Factor", foreignDiscountFactor},
        };
         
         Assert.AreEqual(expected, actual);
    }
    
    [Test]
    public void GarmanKohlhagenSpotOptionPricerInvalidOptionTypeTest()
    {
         const double fxSpotPrice = 1.6;
         const double strike = 1.6;
         const double domesticRiskFreeRate = 0.08;
         const double foreignRiskFreeRate = 0.11;
         const double vol = 0.141;
         const double optionMaturity = 4.0 / 12.0;
         
         string? actual =
             Pricers.GarmanKohlhagenSpotOptionPricer(
                 spotPrice: fxSpotPrice, 
                 strike: strike, 
                 domesticRiskFreeRate: domesticRiskFreeRate, 
                 foreignRiskFreeRate: foreignRiskFreeRate, 
                 vol: vol, 
                 optionMaturity: optionMaturity, 
                 optionType: "Q", 
                 direction: "L", 
                 outputType: "VERBOSE").ToString();
         
         string expected = CommonUtils.DExcelErrorMessage($"Invalid option type: 'Q'");
         Assert.AreEqual(expected, actual);
    }
    
    [Test]
    public void GarmanKohlhagenSpotOptionPricerInvalidDirectionTest()
    {
         const double fxSpotPrice = 1.6;
         const double strike = 1.6;
         const double domesticRiskFreeRate = 0.08;
         const double foreignRiskFreeRate = 0.11;
         const double vol = 0.141;
         const double optionMaturity = 4.0 / 12.0;
         
         string? actual =
             Pricers.GarmanKohlhagenSpotOptionPricer(
                 spotPrice: fxSpotPrice, 
                 strike: strike, 
                 domesticRiskFreeRate: domesticRiskFreeRate, 
                 foreignRiskFreeRate: foreignRiskFreeRate, 
                 vol: vol, 
                 optionMaturity: optionMaturity, 
                 optionType: "C", 
                 direction: "Q").ToString();
         
         string expected = CommonUtils.DExcelErrorMessage($"Invalid direction: 'Q'");
         Assert.AreEqual(expected, actual);
    }
    
    [Test]
    public void GarmanKohlhagenSpotOptionPricerInvalidSpotPriceTest()
    {
         const double fxSpotPrice = -1.6;
         const double strike = 1.6;
         const double domesticRiskFreeRate = 0.08;
         const double foreignRiskFreeRate = 0.11;
         const double vol = 0.141;
         const double optionMaturity = 4.0 / 12.0;
         
         string? actual =
             Pricers.GarmanKohlhagenSpotOptionPricer(
                 spotPrice: fxSpotPrice, 
                 strike: strike, 
                 domesticRiskFreeRate: domesticRiskFreeRate, 
                 foreignRiskFreeRate: foreignRiskFreeRate, 
                 vol: vol, 
                 optionMaturity: optionMaturity, 
                 optionType: "C", 
                 direction: "L").ToString();
         
         string expected = CommonUtils.DExcelErrorMessage($"FX spot price non-positive: {fxSpotPrice}");
         Assert.AreEqual(expected, actual);
    }
    
    [Test]
    public void GarmanKohlhagenSpotOptionPricerInvalidVolTest()
    {
         const double fxSpotPrice = 1.6;
         const double strike = 1.6;
         const double domesticRiskFreeRate = 0.08;
         const double foreignRiskFreeRate = 0.11;
         const double vol = -0.141;
         const double optionMaturity = 4.0 / 12.0;
         
         string? actual =
             Pricers.GarmanKohlhagenSpotOptionPricer(
                 spotPrice: fxSpotPrice, 
                 strike: strike, 
                 domesticRiskFreeRate: domesticRiskFreeRate, 
                 foreignRiskFreeRate: foreignRiskFreeRate, 
                 vol: vol, 
                 optionMaturity: optionMaturity, 
                 optionType: "C", 
                 direction: "L").ToString();
         
         string expected = CommonUtils.DExcelErrorMessage($"Volatility non-positive: {vol}");
         Assert.AreEqual(expected, actual);
    }
}
