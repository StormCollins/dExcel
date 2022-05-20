namespace dExcel;

using System;
using System.IO;
using ExcelDna.Integration;
using mnd = MathNet.Numerics.Distributions;

public static class DerivativeUtils
{
    [ExcelFunction(
       Name = "d.Black",
       Description = "Black option pricer." +
                     "\nShould be used to price options on: futures, forwards, and zero coupon bonds." +
                     "\nTo price swaptions multiply by the relevant annuity factor." +
                     "\nDeprecates AQS function: 'Black'",
       Category = "∂Excel: Derivatives")]
    // TODO: Replace OptionType with an Enum.
    public static object Black(
        [ExcelArgument(Name = "Forward Rate", Description = "Forward rate.")]
        double forwardRate,
        [ExcelArgument(Name = "Risk free rate - NACC", Description = "Risk free rate. Only required for discounting.")]
        double rate,
        [ExcelArgument(Name = "Strike", Description = "Strike.")]
        double strike,
        [ExcelArgument(Name = "Vol", Description = "Vol.")]
        double vol,
        [ExcelArgument(Name = "Option Maturity", Description = "Option maturity.")]
        double optionMaturity,
        [ExcelArgument(Name = "OptionType", Description = "Call or Put - input c or p.")]
        string optionType)
    {
#if DEBUG
        CommonUtils.InFunctionWizard();
#endif 

        int sign = 0;
        
        if (optionType.ToUpper() == "C" || optionType.ToUpper() == "CALL")
        {
            sign = 1;
        }
        else if (optionType.ToUpper() == "P" || optionType.ToUpper() == "PUT")
        {
            sign = -1;
        }
        else
        {
            Console.WriteLine("Error: Invalid option type.");
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
       Name = "d.BlackScholes",
       Description = "Black-Scholes option pricer. \nDeprecates AQS function: 'BS'",
       Category = "∂Excel: Derivatives")]
    public static object BlackScholes(
        [ExcelArgument(Name = "Option Type", Description = "Call or Put - input c or p.")]
        string optionType,
       [ExcelArgument(Name = "S", Description = "Current stock price.")]
        double spotPrice,
        [ExcelArgument(Name = "K", Description = "Strike.")]
        double strike,
       [ExcelArgument(Name = "r", Description = "Risk free rate. Only required for discounting.")]
        double rate,
       [ExcelArgument(Name = "q", Description = "Dividend Yield (NACC).")]
        double dividendYield,
        [ExcelArgument(Name = "T", Description = "Time to maturity.")]
        double timeToMaturity,
        [ExcelArgument(Name = "σ", Description = "Volatility.")]
        double vol)
    {
#if DEBUG
        CommonUtils.InFunctionWizard();
#endif 

        int sign = 0;

        if (optionType.ToUpper() == "C" || optionType.ToUpper() == "CALL")
        {
            sign = 1;
        }
        else if (optionType.ToUpper() == "P" || optionType.ToUpper() == "PUT")
        {
            sign = -1;
        }
        else
        {
            Console.WriteLine("Error: Invalid option type.");
        }

        double d1 = (Math.Log(spotPrice / strike) + (rate-dividendYield+Math.Pow(vol, 2)/2) * timeToMaturity) / (vol * Math.Sqrt(timeToMaturity));
        double d2 = d1 - vol * Math.Sqrt(timeToMaturity);

        if (timeToMaturity <= 0)
        {
            return Math.Exp(-(rate-dividendYield) * timeToMaturity) * Math.Max(0, sign * (spotPrice - strike));
        }

        return sign  * ( spotPrice * Math.Exp(-dividendYield * timeToMaturity) * mnd.Normal.CDF(0, 1, sign * d1) - strike * Math.Exp(-rate * timeToMaturity) * mnd.Normal.CDF(0, 1, sign * d2));
    }

    [ExcelFunction(
       Name = "d.CapletFloorletPrice",
       Description = "Caplet and floorlet pricer. \nDeprecates AQS function: 'CapletFloorletPricer'",
       Category = "∂Excel: Derivatives")]
    public static object CapletFloorlet(
        [ExcelArgument(
            Name = "Forward (Simple)",
            Description = "Provide the simple forward rate.")]
        double forward,
       [ExcelArgument(
            Name = "Rate(NACC)",
            Description = "The NACC interest rate.")]
        double rate,
        [ExcelArgument(
            Name = "Strike (Cap Rate)",
            Description = "The Cap Strike Rate.")]
        double strike,
       [ExcelArgument(
            Name = "Vol",
            Description = "Volatility as at start time.")]
        double vol,
        [ExcelArgument(
            Name = "Start Time",
            Description = "Time until start of option = option maturity")]
        double startTime,
        [ExcelArgument(
            Name = "End Time",
            Description = "Time until end of option = payoff time.")]
        double endTime,
        [ExcelArgument(
            Name = "Option Type",
            Description = "Option is a caplet or a floorlet")]
        string optionType)
    {
#if DEBUG
        CommonUtils.InFunctionWizard();
#endif 

        int sign = 0;

        if (optionType.ToUpper() == "CAPLET")
        {
            sign = 1;
        }
        else if (optionType.ToUpper() == "FLOORLET")
        {
            sign = -1;
        }
        else
        {
            Console.WriteLine("Error: Invalid option type.");
        }

        double d1 = (Math.Log(forward / strike) + Math.Pow(vol, 2) / 2 * startTime) / (vol * Math.Sqrt(startTime));
        double d2 = d1 - vol * Math.Sqrt(startTime);

        if (startTime <= 0)
        {
            return Math.Exp(-rate * startTime) * Math.Max(0, sign * (forward - strike));
        }

        return (endTime-startTime) * sign * Math.Exp(-rate * endTime) * (forward * mnd.Normal.CDF(0, 1, sign * d1) - strike * mnd.Normal.CDF(0, 1, sign * d2));
    }

    [ExcelFunction(
       Name = "d.Bachelier",
       Description = "Bachelier option pricer on Sport or Forwards/Futures." +
                     "\nTo price swaptions you need to multiply by the relevant annuity factor.",
       Category = "∂Excel: Derivatives")]
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
        [ExcelArgument(Name = "Option Type", Description = "Call or Put - input c or p.")]
        string optionType,
        [ExcelArgument(Name = "Forward/Spot", Description = "Specify if it's an option on a forward or spot - f or s.")]
        string forwardOrSpot)
    {
#if DEBUG
        CommonUtils.InFunctionWizard();
#endif 

        int sign = 0;

        if (optionType.ToUpper() == "C" || optionType.ToUpper() == "CALL")
        {
            sign = 1;
        }
        else if (optionType.ToUpper() == "P" || optionType.ToUpper() == "PUT")
        {
            sign = -1;
        }
        else
        {
            Console.WriteLine(@"Error: Invalid option type.");
        }

        double d = (forwardRate - strike) / (vol * Math.Sqrt(optionMaturity));

        var value = sign * (forwardRate - strike) * mnd.Normal.CDF(0, 1, sign * d) + vol * Math.Sqrt(optionMaturity) * mnd.Normal.PDF(0, 1, d);
        return forwardOrSpot.ToUpper() == "F" ? value : Math.Exp(-rate * optionMaturity) * value;
    }
}
