namespace dExcel.InterestRates;

using System;
using mnd = MathNet.Numerics.Distributions;
using ExcelDna.Integration;
using Utilities;

public static class Pricers
{
    /// <summary>
    /// Black pricer for options on Interest Rate Futures and FRAS.
    /// To price swaptions multiply by the relevant annuity factor.
    /// </summary>
    /// <param name="forwardRate">Forward rate for interest rate future/FRA.</param>
    /// <param name="strike">Strike.</param>
    /// <param name="riskFreeRate">Risk free rate (NACC). Only used for discounting.</param>
    /// <param name="vol">Volatility.</param>
    /// <param name="optionMaturity">Option maturity in years.</param>
    /// <param name="optionType">'Call'/'C' or 'Put'/'P'.</param>
    /// <param name="outputType">'VERBOSE' or 'PRICE'. Output full calculation details with 'VERBOSE' or just the price
    /// with 'PRICE'.</param>
    /// <returns>Black price for an option on an interest rate future or FRA.</returns>
    [ExcelFunction(
       Name = "d.IR_BlackForwardOptionPricer",
       Description = 
           "Black pricer for options on Interest Rate Futures and FRAS.\n" +
           "To price swaptions multiply by the relevant annuity factor.\n" +
           "Deprecates AQS function: 'Black'",
       Category = "∂Excel: Interest Rates")]
    // TODO: Replace OptionType with an Enum.
    public static object BlackForwardOptionPricer(
        [ExcelArgument(Name = "Forward Rate", Description = "Forward rate for interest rate future/FRA.")]
        double forwardRate,
        [ExcelArgument(Name = "Strike", Description = "Strike.")]
        double strike,
        [ExcelArgument(Name = "Risk free rate (NACC)",
            Description = "Risk free rate (NACC). Only used for discounting.")]
        double riskFreeRate,
        [ExcelArgument(Name = "Vol", Description = "Volatility.")]
        double vol,
        [ExcelArgument(Name = "Option Maturity", Description = "Option maturity in years.")]
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
        (int? optionTypeSign, string? optionTypeErrorMessage) = CommonUtils.GetSignOfOptionType(optionType);
        if (optionTypeErrorMessage is not null)
        {
            return optionTypeErrorMessage;
        }  
       
        (int? directionSign, string? directionErrorMessage) = CommonUtils.GetSignOfDirection(direction);
        if (directionErrorMessage is not null)
        {
            return directionErrorMessage;
        }
        
        double d1 = (Math.Log(forwardRate / strike) + 0.5 * Math.Pow(vol, 2) * optionMaturity) / (vol * Math.Sqrt(optionMaturity));
        double d2 = d1 - vol * Math.Sqrt(optionMaturity);
        double discountFactor = Math.Exp(-1 * riskFreeRate * optionMaturity);
        double price = 
            (double)directionSign * (double)optionTypeSign * discountFactor * 
            (forwardRate * mnd.Normal.CDF(0, 1, (double) optionTypeSign * d1) - strike * mnd.Normal.CDF(0, 1, (double) optionTypeSign * d2));
        
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

    [ExcelFunction(
        Name = "d.IR_BachelierForwardOptionPricer",
        Description = "Bachelier option pricer on spot or futures/forwards." +
                      "\nTo price swaptions you need to multiply by the relevant annuity factor.",
        Category = "∂Excel: Interest Rates")]
    public static object Bachelier(
        [ExcelArgument(Name = "Forward Rate", Description = "Forward rate.")]
        double forwardRate,
        [ExcelArgument(Name = "Risk Free Rate (NACC)", Description = "Risk free rate. Only required pricing options on forwards.")]
        double rate,
        [ExcelArgument(Name = "Strike", Description = "Strike.")]
        double strike,
        [ExcelArgument(Name = "Vol", Description = "Volatility.")]
        double vol,
        [ExcelArgument(Name = "Option Maturity", Description = "Option maturity.")]
        double optionMaturity,
        [ExcelArgument(Name = "Option Type", Description = "'Call'/'C' or 'Put'/'P'.")]
        string optionType,
        [ExcelArgument(Name = "Forward/Spot", Description = "Option on 'forward' or 'spot'.")]
        string forwardOrSpot,
        [ExcelArgument(Name = "Long/Short", Description = "'Long' or 'Short'.")]
        string longOrShort)
    {
#if DEBUG
        CommonUtils.InFunctionWizard();
#endif 
        (int? sign, string? errorMessage) = CommonUtils.GetSignOfOptionType(optionType);
        if (errorMessage is not null)
        {
            return errorMessage;
        }  

        double d = (forwardRate - strike) / (vol * Math.Sqrt(optionMaturity));
        double value = (double)sign * (forwardRate - strike) * mnd.Normal.CDF(0, 1, (double)sign * d) + vol * Math.Sqrt(optionMaturity) * mnd.Normal.PDF(0, 1, d);
        int longOrShortDirection;
        switch (longOrShort.ToUpper())
        {
            case "LONG":
                longOrShortDirection = 1;
                break;
            case "SHORT":
                longOrShortDirection = -1;
                break;
            default:
                return CommonUtils.DExcelErrorMessage($"Invalid 'long'/'short' direction: {longOrShort}");
        }

        return longOrShortDirection * (forwardOrSpot.ToUpper() == "F" ? value : Math.Exp(-rate * optionMaturity) * value);
    }
}
