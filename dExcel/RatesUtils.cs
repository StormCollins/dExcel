namespace dExcel;

using System;
using ExcelDna.Integration;

public static class RatesUtils
{
    [ExcelFunction(
       Name = "d.Disc2ForwardRate",
       Description = "Calculates the forward rate from two discount factors. \n" +
                     "Deprecates AQS Function: 'Disc2ForwardRate'",
       Category = "∂Excel: Interest Rates")]
    public static object Disc2ForwardRate(
       [ExcelArgument(
            Name = "RateType",
            Description = "Type of Rate: 'simple' = 'simple', 'naca' = 'naca', 'nacs' = 'nacs', 'nacq' = 'nacq', 'nacm' = 'nacm', 'nacc' = 'nacc' ")]
        string rateType,
       [ExcelArgument(
            Name = "DFPrevious",
            Description = "Discount Factor at time t-1.")]
        double dFPrevious,
       [ExcelArgument(
            Name = "DFCurrent",
            Description = "Discount Factor at time t.")]
        double dFCurrent,
       [ExcelArgument(
            Name = "dT",
            Description = "Change in time.")]
        double dT)
    {
#if DEBUG
        CommonUtils.InFunctionWizard();
#endif 

        object getForward = 0;
        switch (rateType.ToUpper())
        {
            case "SIMPLE":
                getForward = (dFPrevious / dFCurrent - 1) / dT;
                break;
            case "NACC":
                getForward = (Math.Log(dFPrevious) - Math.Log(dFCurrent)) / dT;
                break;
            case "NACA":
                getForward = Math.Exp((Math.Log(dFPrevious) - Math.Log(dFCurrent)) / dT) - 1;
                break;
            case "NACS":
                getForward = (Math.Pow(Math.Exp((Math.Log(dFPrevious) - Math.Log(dFCurrent)) / dT), 1.0 / 2.0) - 1) * 2;
                break;
            case "NACQ":
                getForward = (Math.Pow(Math.Exp((Math.Log(dFPrevious) - Math.Log(dFCurrent)) / dT), 1.0 / 4.0) - 1) * 4;
                break;
            case "NACM":
                getForward = (Math.Pow(Math.Exp((Math.Log(dFPrevious) - Math.Log(dFCurrent)) / dT), 1.0 / 12.0) - 1) * 12;
                break;
            default:
                break;
        }
        return getForward;
    }

    [ExcelFunction(
       Name = "d.IntConvert",
       Description = "Calculates the forward rate from two discount factors. \n" +
                     "Deprecates AQS Function: 'IntConvert'",
       Category = "∂Excel: Interest Rates")]
    public static object IntConvert(
        [ExcelArgument(
                Name = "Rate",
                Description = "Rate we want to convert.")]
        double rate,
        [ExcelArgument(
                Name = "RateFrom",
                Description = "Rate we want to convert from.")]
        string RateFrom,
        [ExcelArgument(
                Name = "RateTo",
                Description = "Rate we want to convert to.")]
        string RateTo)
    {
#if DEBUG
        CommonUtils.InFunctionWizard();
#endif

        // First we do all conversions from simple rate to other.
        if (RateFrom.ToUpper() == "SIMPLE" && RateTo.ToUpper() == "NACA")
        {
            return (Math.Pow(1 + rate, 1) - 1);
        }
        else if (RateFrom.ToUpper() == "SIMPLE" && RateTo.ToUpper() == "NACS")
        {
            return 2 * (Math.Pow(1 + rate, 1.0 / 2.0) - 1);
        }
        else if (RateFrom.ToUpper() == "SIMPLE" && RateTo.ToUpper() == "NACQ")
        {
            return 4 * (Math.Pow(1 + rate, 1.0 / 4.0) - 1);
        }
        else if (RateFrom.ToUpper() == "SIMPLE" && RateTo.ToUpper() == "NACM")
        {
            return 12 * (Math.Pow(1 + rate, 1.0 / 12.0) - 1);
        }
        else if (RateFrom.ToUpper() == "SIMPLE" && RateTo.ToUpper() == "NACC")
        {
            return Math.Log(1+rate);
        }
        // Second we do all conversions from rates to simple.
        else if (RateFrom.ToUpper() == "NACA" && RateTo.ToUpper() == "SIMPLE")
        {
            return Math.Pow(1 + rate / 1.0, 1) - 1;
        }
        else if (RateFrom.ToUpper() == "NACS" && RateTo.ToUpper() == "SIMPLE")
        {
            return Math.Pow(1 + rate / 2.0, 2) - 1;
        }
        else if (RateFrom.ToUpper() == "NACQ" && RateTo.ToUpper() == "SIMPLE")
        {
            return Math.Pow(1 + rate / 4.0, 4) - 1;
        }
        else if (RateFrom.ToUpper() == "NACM" && RateTo.ToUpper() == "SIMPLE")
        {
            return Math.Pow(1 + rate / 12.0, 12) - 1;
        }
        else if (RateFrom.ToUpper() == "NACC" && RateTo.ToUpper() == "SIMPLE")
        {
            return Math.Exp(rate) - 1;
        }
        // Third we do all conversions from NACx rate to NACy (excluding NACC).
        else if (RateFrom.ToUpper() == "NACA" && RateTo.ToUpper() == "NACS")
        {
            return 2*(Math.Pow(1+rate,1/2.0)-1);
        }
        else if (RateFrom.ToUpper() == "NACA" && RateTo.ToUpper() == "NACQ")
        {
            return 4 * (Math.Pow(1 + rate, 1 / 4.0) - 1);
        }
        else if (RateFrom.ToUpper() == "NACA" && RateTo.ToUpper() == "NACM")
        {
            return 12 * (Math.Pow(1 + rate, 1 / 12.0) - 1);
        }
        else if (RateFrom.ToUpper() == "NACS" && RateTo.ToUpper() == "NACA")
        {
            return 1 * (Math.Pow(1 + rate/2, 2 / 1) - 1);
        }
        else if (RateFrom.ToUpper() == "NACS" && RateTo.ToUpper() == "NACQ")
        {
            return 4 * (Math.Pow(1 + rate/2, 2 / 4.0) - 1);
        }
        else if (RateFrom.ToUpper() == "NACS" && RateTo.ToUpper() == "NACM")
        {
            return 12 * (Math.Pow(1 + rate/2, 2 / 12.0) - 1);
        }
        else if (RateFrom.ToUpper() == "NACQ" && RateTo.ToUpper() == "NACA")
        {
            return 1 * (Math.Pow(1 + rate/4, 4) - 1);
        }
        else if (RateFrom.ToUpper() == "NACQ" && RateTo.ToUpper() == "NACS")
        {
            return 2 * (Math.Pow(1 + rate / 4, 4/2) - 1);
        }
        else if (RateFrom.ToUpper() == "NACQ" && RateTo.ToUpper() == "NACM")
        {
            return 12 * (Math.Pow(1 + rate / 4, 4 / 12.0) - 1);
        }
        else if (RateFrom.ToUpper() == "NACM" && RateTo.ToUpper() == "NACA")
        {
            return 1 * (Math.Pow(1 + rate / 12, 12 / 1) - 1);
        }
        else if (RateFrom.ToUpper() == "NACM" && RateTo.ToUpper() == "NACS")
        {
            return 2 * (Math.Pow(1 + rate / 12, 12 / 2) - 1);
        }
        else if (RateFrom.ToUpper() == "NACM" && RateTo.ToUpper() == "NACQ")
        {
            return 4 * (Math.Pow(1 + rate / 12, 12 / 4) - 1);
        }
        // Fourth we do all conversions from NACx rate to NACC.
        else if (RateFrom.ToUpper() == "NACA" && RateTo.ToUpper() == "NACC")
        {
            return 1 * Math.Log(1 + rate / 1) ;
        }
        else if (RateFrom.ToUpper() == "NACS" && RateTo.ToUpper() == "NACC")
        {
            return 2 * Math.Log(1 + rate / 2);
        }
        else if (RateFrom.ToUpper() == "NACQ" && RateTo.ToUpper() == "NACC")
        {
            return 4 * Math.Log(1 + rate / 4);
        }
        else if (RateFrom.ToUpper() == "NACM" && RateTo.ToUpper() == "NACC")
        {
            return 12 * Math.Log(1 + rate / 12);
        }
        // Fifth we do all conversions from NACC rate to NACx.
        else if (RateFrom.ToUpper() == "NACC" && RateTo.ToUpper() == "NACA")
        {
            return 1 * (Math.Exp(rate / 1)-1);
        }
        else if (RateFrom.ToUpper() == "NACC" && RateTo.ToUpper() == "NACS")
        {
            return 2 * (Math.Exp(rate / 2) - 1);
        }
        else if (RateFrom.ToUpper() == "NACC" && RateTo.ToUpper() == "NACQ")
        {
            return 4 * (Math.Exp(rate / 4) - 1);
        }
        else if (RateFrom.ToUpper() == "NACC" && RateTo.ToUpper() == "NACM")
        {
            return 12 * (Math.Exp(rate / 12) - 1);
        }
        // Sixth - if the same rate is used for conversion, it should output that rate
        else if (RateFrom.ToUpper() == RateTo.ToUpper() )
        {
            return rate;
        }
        else
        {
            return "Invalid Rate";
        }
    }
}
