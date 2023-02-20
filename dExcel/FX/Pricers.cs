namespace dExcel.FX;

using ExcelDna.Integration;
using Utilities;
using mnd = MathNet.Numerics.Distributions;

/// <summary>
/// A collection of pricers for equity derivatives.
/// </summary>
public static class Pricers
{
    /// <summary>
    /// Garman-Kohlhagen pricer for an option on foreign exchange (FX) spot.
    /// </summary>
    /// <param name="spotPrice">Spot price.</param>
    /// <param name="strike">Strike.</param>
    /// <param name="domesticRiskFreeRate">Domestic risk free rate (NACC). Used for discounting.</param>
    /// <param name="foreignRiskFreeRate">Foreign risk free rate (NACC). Used for forecasting.</param>
    /// <param name="optionMaturity">Option maturity in years.</param>
    /// <param name="vol">Volatility.</param>
    /// <param name="optionType">'Call'/'C' or 'Put'/'P'.</param>
    /// <param name="direction">'Long'/'L'/'Buy'/'B' or 'Short'/'S'/'Sell'.</param>
    /// <param name="outputType">'VERBOSE' or 'PRICE'. Output full calculation details with 'VERBOSE' or just the price
    /// with 'PRICE'.</param>
    /// <returns>Garman-Kohlhagen price for an option on foreign exchange (FX) spot.</returns>
    [ExcelFunction(
       Name = "d.FX_GarmanKohlhagenFXOptionPricer",
       Description = "Garman-Kohlhagen pricer for an option on FX spot.",
       Category = "∂Excel: FX")]
    public static object GarmanKohlhagenSpotOptionPricer(
        [ExcelArgument(Name = "X₀", Description = "Initial FX spot price.")]
        double spotPrice,
        [ExcelArgument(Name = "K", Description = "Strike.")]
        double strike,
        [ExcelArgument(Name = "domesticRiskFreeRate", Description = "Domestic risk free rate (NACC).")]
        double domesticRiskFreeRate,
        [ExcelArgument(Name = "foreignRiskFreeRate", Description = "Foreign risk free rate (NACC).")]
        double foreignRiskFreeRate,
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
            return CommonUtils.DExcelErrorMessage($"FX spot price non-positive: {spotPrice}");
        }

        if (vol <= 0)
        {
            return CommonUtils.DExcelErrorMessage($"Volatility non-positive: {vol}");
        }

        double d1 = (Math.Log(spotPrice / strike) + (domesticRiskFreeRate - foreignRiskFreeRate + Math.Pow(vol, 2)/2) * optionMaturity) / (vol * Math.Sqrt(optionMaturity));
        double d2 = d1 - vol * Math.Sqrt(optionMaturity);
        double domesticDiscountFactor = Math.Exp(-1 * domesticRiskFreeRate * optionMaturity);
        double foreignDiscountFactor = Math.Exp(-1 * foreignRiskFreeRate * optionMaturity);
        double price = 
            (double)directionSign * ((double)optionTypeSign * 
                (spotPrice * foreignDiscountFactor * mnd.Normal.CDF(0, 1, (double)optionTypeSign * d1) -
                        strike * domesticDiscountFactor * mnd.Normal.CDF(0, 1, (double)optionTypeSign * d2)));
        
        if (outputType.ToUpper() == "PRICE")
        {
            return price;
        }

        object[,] results = 
        {
            {"Price", price},
            {"d1", d1},
            {"d2", d2},
            {"Domestic Discount Factor", domesticDiscountFactor},
            {"Foreign Discount Factor", foreignDiscountFactor},
        };

        return results;
    }
}
