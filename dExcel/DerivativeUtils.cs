namespace dExcel;

using System;
using ExcelDna.Integration;
using mnd = MathNet.Numerics.Distributions;

public static class DerivativeUtils
{
 
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

   }
