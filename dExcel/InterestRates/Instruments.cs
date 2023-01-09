using System.Security.Cryptography;

namespace dExcel.InterestRates;

using System;
using mnd = MathNet.Numerics.Distributions;
using ExcelDna.Integration;
using QLNet;
using Utilities;

public static class Instruments
{
    [ExcelFunction(
       Name = "d.IR_Black",
       Description = "Black option pricer." +
                 "\nShould be used to price options on: Interest rate futures and FRAs" +
                 "\nTo price swaptions multiply by the relevant annuity factor." +
                 "\nDeprecates AQS function: 'Black'",
       Category = "∂Excel: Interest Rates")]
    // TODO: Replace OptionType with an Enum.
    public static object Black(
        [ExcelArgument(Name = "Forward Rate", Description = "Forward rate.")]
        double forwardRate,
        [ExcelArgument(Name = "Risk free rate(NACC)", Description = "Risk free rate. Only required for discounting.")]
        double rate,
        [ExcelArgument(Name = "Strike", Description = "Strike.")]
        double strike,
        [ExcelArgument(Name = "Vol", Description = "Volatility.")]
        double vol,
        [ExcelArgument(Name = "Option Maturity", Description = "Option maturity.")]
        double optionMaturity,
        [ExcelArgument(Name = "OptionType", Description = "Call/C or Put/P.")]
        string optionType)
    {
#if DEBUG
        CommonUtils.InFunctionWizard();
#endif 
        int sign = 0;
        switch (optionType.ToUpper())
        {
            case "C":
            case "CALL":
                sign = 1;
                break;
            case "P":
            case "PUT":
                sign = -1;
                break;
            default:
                return CommonUtils.DExcelErrorMessage("Invalid option type.");
        }

        double d1 = (Math.Log(forwardRate / strike) + Math.Pow(vol, 2) * optionMaturity / 2) / (vol * Math.Sqrt(optionMaturity));
        double d2 = d1 - vol * Math.Sqrt(optionMaturity);

        if (optionMaturity <= 0)
        {
            return Math.Exp(-rate * optionMaturity) * Math.Max(0, sign * (forwardRate - strike));
        }
        
        return sign * Math.Exp(-rate * optionMaturity) * (forwardRate * mnd.Normal.CDF(0, 1, sign * d1) - strike * mnd.Normal.CDF(0, 1, sign * d2));            
    }

    [ExcelFunction(
        Name = "d.IR_Bachelier",
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
        int sign;
        switch (optionType.ToUpper())
        {
            case "C":
            case "CALL":
                sign = 1;
                break;
            case "P":
            case "PUT":
                sign = -1;
                break;
            default:
                return CommonUtils.DExcelErrorMessage($"Invalid option type: {optionType}");
        }

        double d = (forwardRate - strike) / (vol * Math.Sqrt(optionMaturity));
        double value = sign * (forwardRate - strike) * mnd.Normal.CDF(0, 1, sign * d) + vol * Math.Sqrt(optionMaturity) * mnd.Normal.PDF(0, 1, d);
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
