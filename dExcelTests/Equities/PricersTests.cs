namespace dExcelTests.Equities;

using dExcel.Equities;
using dExcel.Utilities;
using NUnit.Framework;
using mnd = MathNet.Numerics.Distributions;
using QLNet;

[TestFixture]
public class PricersTests
{
    [Test]
    public void BlackScholesSpotOptionPricerForCallOptionTest()
    {
         // See Example 15.6 of John Hull - Options, Futures, and Other Derivatives, 9th Edition. 
         // Call option price = 4.76
         const double initialSpot = 42;
         const double strike = 40;
         const double riskFreeRate = 0.1;
         const double dividendYield = 0.0;
         const double vol = 0.2;
         const double optionMaturity = 0.5;

         double actualCallPrice = 
             (double)Pricers.BlackScholesSpotOptionPricer(
                 spotPrice: initialSpot, 
                 strike: strike, 
                 riskFreeRate: riskFreeRate, 
                 dividendYield: dividendYield,
                 vol: vol, 
                 optionMaturity: optionMaturity, 
                 optionType: "C", 
                 direction: "L");

         double d1 = (Math.Log(initialSpot / strike) + (riskFreeRate + 0.5 * vol * vol) * optionMaturity) /
                     (vol * Math.Sqrt(optionMaturity));
         
         double d2 = d1 - vol * Math.Sqrt(optionMaturity);

         double expectedCallPrice = 
             initialSpot * mnd.Normal.CDF(0, 1, d1) - 
             strike * Math.Exp(-1 * riskFreeRate * optionMaturity) * mnd.Normal.CDF(0, 1, d2);
         
         Assert.AreEqual(expectedCallPrice, actualCallPrice, 1e-6);
    }

    [Test]
    public void BlackScholesOptionPricerForPutOptionTest()
    {
         // See Example 15.6 of John Hull - Options, Futures, and Other Derivatives, 9th Edition. 
         // Put option price = 0.81.
         const double initialSpot = 42;
         const double strike = 40;
         const double riskFreeRate = 0.1;
         const double dividendYield = 0.0;
         const double vol = 0.2;
         const double optionMaturity = 0.5;

         double actualPutPrice = 
             (double)Pricers.BlackScholesSpotOptionPricer(
                 spotPrice: initialSpot, 
                 strike: strike, 
                 riskFreeRate: riskFreeRate, 
                 dividendYield: dividendYield,
                 vol: vol, 
                 optionMaturity: optionMaturity, 
                 optionType: "P", 
                 direction: "L");

         double d1 = (Math.Log(initialSpot / strike) + (riskFreeRate + 0.5 * vol * vol) * optionMaturity) /
                     (vol * Math.Sqrt(optionMaturity));
         
         double d2 = d1 - vol * Math.Sqrt(optionMaturity);

         double expectedPutPrice = 
             -1 * (initialSpot * mnd.Normal.CDF(0, 1, -1 * d1) - 
             strike * Math.Exp(-1 * riskFreeRate * optionMaturity) * mnd.Normal.CDF(0, 1, -1 * d2));
         
         Assert.AreEqual(expectedPutPrice, actualPutPrice, 1e-6);
    }

    [Test]
    public void BlackScholesSpotOptionPricerForCallOptionVerboseOutputTest()
    {
         // See Example 15.6 of John Hull - Options, Futures, and Other Derivatives, 9th Edition. 
         // Call option price = 4.76
         const double initialSpot = 42;
         const double strike = 40;
         const double riskFreeRate = 0.1;
         const double dividendYield = 0.0;
         const double vol = 0.2;
         const double optionMaturity = 0.5;

         object[,] actual = 
             (object[,])Pricers.BlackScholesSpotOptionPricer(
                 spotPrice: initialSpot, 
                 strike: strike, 
                 riskFreeRate: riskFreeRate, 
                 dividendYield: dividendYield,
                 vol: vol, 
                 optionMaturity: optionMaturity, 
                 optionType: "C", 
                 direction: "L",
                 outputType: "VERBOSE");

         double d1 = (Math.Log(initialSpot / strike) + (riskFreeRate + 0.5 * vol * vol) * optionMaturity) /
                     (vol * Math.Sqrt(optionMaturity));
         
         double d2 = d1 - vol * Math.Sqrt(optionMaturity);
         double discountFactor = Math.Exp(-1 * riskFreeRate * optionMaturity);
         double price = initialSpot * mnd.Normal.CDF(0, 1, d1) - strike * discountFactor * mnd.Normal.CDF(0, 1, d2);
         
         object[,] expected = 
         {
             {"Price", price},
             {"d1", d1},
             {"d2", d2},
             {"Discount Factor", discountFactor},
         };
         
         Assert.AreEqual(expected, actual);
    }
    
    [Test]
    public void BlackScholesOptionPricerInvalidOptionTypeTest()
    {
         const double initialSpot = 42;
         const double strike = 40;
         const double riskFreeRate = 0.1;
         const double dividendYield = 0.0;
         const double vol = 0.2;
         const double optionMaturity = 0.5;

         string actual = 
             Pricers.BlackScholesSpotOptionPricer(
                 spotPrice: initialSpot, 
                 strike: strike, 
                 riskFreeRate: riskFreeRate, 
                 dividendYield: dividendYield,
                 vol: vol, 
                 optionMaturity: optionMaturity, 
                 optionType: "Q", 
                 direction: "L").ToString();
         
         string expected = CommonUtils.DExcelErrorMessage($"Invalid option type: 'Q'");
         Assert.AreEqual(expected, actual);
    }
    
    [Test]
    public void BlackScholesOptionPricerInvalidDirectionTest()
    {
         const double initialSpot = 42;
         const double strike = 40;
         const double riskFreeRate = 0.1;
         const double dividendYield = 0.0;
         const double vol = 0.2;
         const double optionMaturity = 0.5;

         string actual = 
             Pricers.BlackScholesSpotOptionPricer(
                 spotPrice: initialSpot, 
                 strike: strike, 
                 riskFreeRate: riskFreeRate, 
                 dividendYield: dividendYield,
                 vol: vol, 
                 optionMaturity: optionMaturity, 
                 optionType: "P", 
                 direction: "Q").ToString();
         
         string expected = CommonUtils.DExcelErrorMessage($"Invalid direction: 'Q'");
         Assert.AreEqual(expected, actual);
    }
    
    [Test]
    public void BlackScholesOptionPricerInvalidSpotPriceTest()
    {
         const double initialSpot = -42;
         const double strike = 40;
         const double riskFreeRate = 0.1;
         const double dividendYield = 0.0;
         const double vol = 0.2;
         const double optionMaturity = 0.5;

         string actual = 
             Pricers.BlackScholesSpotOptionPricer(
                 spotPrice: initialSpot, 
                 strike: strike, 
                 riskFreeRate: riskFreeRate, 
                 dividendYield: dividendYield,
                 vol: vol, 
                 optionMaturity: optionMaturity, 
                 optionType: "P", 
                 direction: "L").ToString();
         
         string expected = CommonUtils.DExcelErrorMessage($"Spot price non-positive: {initialSpot}");
         Assert.AreEqual(expected, actual);
    }
    
    [Test]
    public void BlackScholesOptionPricerInvalidVolTest()
    {
         const double initialSpot = 42;
         const double strike = 40;
         const double riskFreeRate = 0.1;
         const double dividendYield = 0.0;
         const double vol = -0.2;
         const double optionMaturity = 0.5;

         string actual = 
             Pricers.BlackScholesSpotOptionPricer(
                 spotPrice: initialSpot, 
                 strike: strike, 
                 riskFreeRate: riskFreeRate, 
                 dividendYield: dividendYield,
                 vol: vol, 
                 optionMaturity: optionMaturity, 
                 optionType: "P", 
                 direction: "L").ToString();
         
         string expected = CommonUtils.DExcelErrorMessage($"Volatility non-positive: {vol}");
         Assert.AreEqual(expected, actual);
    }
    
    [Test]
    public void BlackScholesOptionPricerInvalidDividendYieldTest()
    {
         const double initialSpot = 42;
         const double strike = 40;
         const double riskFreeRate = 0.1;
         const double dividendYield = -0.1;
         const double vol = 0.2;
         const double optionMaturity = 0.5;

         string actual = 
             Pricers.BlackScholesSpotOptionPricer(
                 spotPrice: initialSpot, 
                 strike: strike, 
                 riskFreeRate: riskFreeRate, 
                 dividendYield: dividendYield,
                 vol: vol, 
                 optionMaturity: optionMaturity, 
                 optionType: "P", 
                 direction: "L").ToString();
         
         string expected = CommonUtils.DExcelErrorMessage($"Dividend yield non-positive: {dividendYield}");
         Assert.AreEqual(expected, actual);
    }
}
