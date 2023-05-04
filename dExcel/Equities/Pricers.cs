using dExcel.Utilities;
using ExcelDna.Integration;
using mnd = MathNet.Numerics.Distributions;

namespace dExcel.Equities;

/// <summary>
/// A collection of pricers for equity derivatives.
/// </summary>
public static class Pricers
{
    /// <summary>
    /// Black-Scholes pricer for option on equity spot.
    /// </summary>
    /// <param name="spotPrice">Spot price.</param>
    /// <param name="strike">Strike.</param>
    /// <param name="riskFreeRate">Risk free rate (NACC). Used for discounting.</param>
    /// <param name="dividendYield">Dividend yield.</param>
    /// <param name="optionMaturity">Option maturity in years.</param>
    /// <param name="vol">Volatility.</param>
    /// <param name="optionType">'Call'/'C' or 'Put'/'P'.</param>
    /// <param name="direction">'Long'/'L'/'Buy'/'B' or 'Short'/'S'/'Sell'.</param>
    /// <param name="outputType">'VERBOSE' or 'PRICE'. Output full calculation details with 'VERBOSE' or just the price
    /// with 'PRICE'.</param>
    /// <returns>Black-Scholes price for an option on equity spot.</returns>
    [ExcelFunction(
       Name = "d.Equity_BlackScholesSpotOptionPricer",
       Description = "Black-Scholes pricer for an option on equity spot. \nDeprecates AQS function: 'BS'",
       Category = "∂Excel: Equities")]
    public static object BlackScholesSpotOptionPricer(
        [ExcelArgument(Name = "S₀", Description = "Initial stock price.")]
        double spotPrice,
        [ExcelArgument(Name = "K", Description = "Strike.")]
        double strike,
        [ExcelArgument(Name = "r", Description = "Risk free rate (NACC). Only used for discounting.")]
        double riskFreeRate,
        [ExcelArgument(Name = "q", Description = "Dividend Yield (NACC).")]
        double dividendYield,
        [ExcelArgument(Name = "Vol", Description = "Volatility.")]
        double vol,
        [ExcelArgument(Name = "T", Description = "Option maturity in years.")]
        double optionMaturity,
        [ExcelArgument(Name = "Option Type", Description = "'Call'/'C' or 'Put'/'P'.")]
        string optionType,
        [ExcelArgument(Name = "Direction", Description = "'Long'/'L'/'Buy'/'B' or 'Short'/'S'/'Sell'.")]
        string direction,
        [ExcelArgument(
            Name = "(Optional)Output Type", 
            Description = 
                "'VERBOSE' or 'PRICE'.\n" +
                "Output full calculation details with 'VERBOSE' or just the price with 'PRICE'.\n" +
                "Default = 'PRICE'")]
        string outputType = "PRICE")
    {
#if DEBUG
        CommonUtils.InFunctionWizard();
#endif 
        if (!CommonUtils.TryParseOptionTypeToSign(optionType, out int? optionTypeSign, out string? optionTypeErrorMessage))
        {
            return optionTypeErrorMessage;
        }  
       
        if (!CommonUtils.TryParseDirectionToSign(direction, out int? directionSign, out string? directionErrorMessage))
        {
            return directionErrorMessage;
        }
        
        if (spotPrice <= 0)
        {
            return CommonUtils.DExcelErrorMessage($"Spot price non-positive: {spotPrice}");
        }

        if (vol <= 0)
        {
            return CommonUtils.DExcelErrorMessage($"Volatility non-positive: {vol}");
        }

        if (dividendYield < 0)
        {
            return CommonUtils.DExcelErrorMessage($"Dividend yield non-positive: {dividendYield}");
        }

        double d1 = (Math.Log(spotPrice / strike) + (riskFreeRate - dividendYield + Math.Pow(vol, 2)/2) * optionMaturity) / (vol * Math.Sqrt(optionMaturity));
        double d2 = d1 - vol * Math.Sqrt(optionMaturity);
        double discountFactor = Math.Exp(-1 * riskFreeRate * optionMaturity);
        double price = 
            (double)directionSign * ((double)optionTypeSign * 
                (spotPrice * Math.Exp(-1 * dividendYield * optionMaturity) * mnd.Normal.CDF(0, 1, (double)optionTypeSign * d1) -
                        strike * discountFactor * mnd.Normal.CDF(0, 1, (double)optionTypeSign * d2)));
        
        if (outputType.ToUpper() == "PRICE")
        {
            return price;
        }

        object[,] results = 
        {
            {"Price", price},
            {"d1", d1},
            {"d2", d2},
            {"Discount Factor", discountFactor},
        };

        return results;
    }
}
