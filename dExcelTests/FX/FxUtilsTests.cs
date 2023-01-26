namespace dExcelTests.FX;

using NUnit.Framework;
using dExcel.FX;

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
}
