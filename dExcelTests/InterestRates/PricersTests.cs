namespace dExcelTests.InterestRates;

using dExcel.InterestRates;
using dExcel.Utilities;
using NUnit.Framework;

[TestFixture]
public class PricersTests
{
    [Test]
    public void LongBlackForwardOptionPricerTest()
    {
         // See Example 18.6 of John Hull - Options, Futures, and Other Derivatives, 9th Edition. 
         // The example in question is of a commodity option but the same principles apply.
         const double forwardRate = 20;
         const double strike = 20;
         const double riskFreeRate = 0.09;
         const double vol = 0.25;
         const double optionMaturity = 4.0 / 12.0;
         double price = (double)Pricers.BlackForwardOptionPricer(forwardRate, strike, riskFreeRate, vol, optionMaturity, "P", "L");
         Assert.AreEqual(1.1166414565589438, price);
    }

    [Test]
    public void ShortBlackForwardOptionPricerTest()
    {
         // See Example 18.6 of John Hull - Options, Futures, and Other Derivatives, 9th Edition. 
         // The example in question is of a commodity option but the same principles apply.
         const double forwardRate = 20;
         const double strike = 20;
         const double riskFreeRate = 0.09;
         const double vol = 0.25;
         const double optionMaturity = 4.0 / 12.0;
         double price = (double)Pricers.BlackForwardOptionPricer(forwardRate, strike, riskFreeRate, vol, optionMaturity, "P", "S");
         Assert.AreEqual(-1.1166414565589438, price);
    }
    
    [Test]
    public void BlackForwardOptionPricerVerboseTest()
    {
         // See Example 18.6 of John Hull - Options, Futures, and Other Derivatives, 9th Edition. 
         // The example in question is of a commodity option but the same principles apply.
         const double forwardRate = 20;
         const double strike = 20;
         const double riskFreeRate = 0.09;
         const double vol = 0.25;
         const double optionMaturity = 4.0 / 12.0;
         object[,] actual = (object[,])Pricers.BlackForwardOptionPricer(forwardRate, strike, riskFreeRate, vol, optionMaturity, "P", "L", "VERBOSE");
         object[,] expected =
         {
             { "Price", 1.1166414565589438 },
             { "d1", 0.072168783648703216 },
             { "d2", -0.072168783648703216 },
             { "Discount Factor", Math.Exp(-1 * riskFreeRate * optionMaturity) }, 
         };
         
         Assert.AreEqual(expected, actual);
    }
    
    [Test]
    public void BlackForwardOptionPricerInvalidOptionTypeTest()
    {
         const double forwardRate = 20;
         const double strike = 20;
         const double riskFreeRate = 0.09;
         const double vol = 0.25;
         const double optionMaturity = 4.0 / 12.0;
         string? actual = 
              Pricers.BlackForwardOptionPricer(
                   forwardRate: forwardRate, 
                   strike: strike, 
                   riskFreeRate: riskFreeRate, 
                   vol: vol, 
                   optionMaturity: optionMaturity, 
                   optionType: "Q", 
                   direction: "L").ToString();
         
         Assert.AreEqual(CommonUtils.DExcelErrorMessage("Invalid option type: 'Q'"), actual);
    }

    [Test]
    public void BlackForwardOptionPricerInvalidDirectionTest()
    {
         const double forwardRate = 20;
         const double strike = 20;
         const double riskFreeRate = 0.09;
         const double vol = 0.25;
         const double optionMaturity = 4.0 / 12.0;
         string? actual = 
              Pricers.BlackForwardOptionPricer(
                   forwardRate: forwardRate, 
                   strike: strike, 
                   riskFreeRate: riskFreeRate, 
                   vol: vol, 
                   optionMaturity: optionMaturity, 
                   optionType: "P", 
                   direction: "Q").ToString();
         
         Assert.AreEqual(CommonUtils.DExcelErrorMessage("Invalid direction: 'Q'"), actual);
    }
    
    [Test]
    public void BlackForwardOptionPricerPutCallParityTest()
    {
        // See section 18.10 of John Hull - Options, Futures, and Other Derivatives, 9th Edition. 
        // Put-Call parity for an option on a future is given by:
        // C + K = P + F
        // => C = P + F - K
        const double forwardRate = 0.12;
        const double strike = 0.10;
        const double riskFreeRate = 0.09;
        const double vol = 0.25;
        const double optionMaturity = 4.0 / 12.0;
        double discountFactor = Math.Exp(-1 * riskFreeRate * optionMaturity);
        double putPrice = 1 / discountFactor * (double)Pricers.BlackForwardOptionPricer(forwardRate, strike, riskFreeRate, vol, optionMaturity, "P", "L");
        double callPrice = 1 / discountFactor * (double)Pricers.BlackForwardOptionPricer(forwardRate, strike, riskFreeRate, vol, optionMaturity, "C", "L");
        Assert.AreEqual(callPrice, putPrice + forwardRate - strike, 1e-10);
    }
}
