namespace dExcel.InterestRates;

using ExcelDna.Integration;

public static class RatesUtils
{
    /// <summary>
    /// Calculates the forward rate from two discount factors.
    /// </summary>
    /// <param name="compoundingConvention">The compounding convention.</param>
    /// <param name="nearDiscountFactor">The discount factor nearest to the discount curve base date.</param>
    /// <param name="farDiscountFactor">The discount factor furthest from the discount curve base date.</param>
    /// <param name="dT">The year fraction between discount factors.</param>
    /// <returns>Forward rate over the period dT.</returns>
    [ExcelFunction(
       Name = "d.IR_DiscountFactorsToForwardRate",
       Description = "Calculates the forward rate from two discount factors.\n" +
                     "Deprecates AQS Function: 'Disc2ForwardRate'",
       Category = "∂Excel: Interest Rates")]
    public static object Disc2ForwardRate(
       [ExcelArgument(
            Name = "Compounding Convention",
            Description = "Options: 'Simple', 'NACA', 'NACS', 'NACQ', 'NACM', 'NACC'")]
        string compoundingConvention,
       [ExcelArgument(
            Name = "Near DF",
            Description = "Discount factor nearest to the discount curve base date.")]
        double nearDiscountFactor,
       [ExcelArgument(
            Name = "Far DF",
            Description = "Discount factor furthest from the discount curve base date.")]
        double farDiscountFactor,
       [ExcelArgument(
            Name = "ΔT",
            Description = "The year fraction between discount factors.")]
        double dT)
    {
#if DEBUG
        CommonUtils.InFunctionWizard();
#endif

        return compoundingConvention.ToUpper() switch
        {
            "SIMPLE" => (nearDiscountFactor / farDiscountFactor - 1) / dT,
            "NACC" => (Math.Log(nearDiscountFactor) - Math.Log(farDiscountFactor)) / dT,
            "NACA" => Math.Exp((Math.Log(nearDiscountFactor) - Math.Log(farDiscountFactor)) / dT) - 1,
            "NACS" => (Math.Pow(Math.Exp((Math.Log(nearDiscountFactor) - Math.Log(farDiscountFactor)) / dT), 1.0 / 2.0) - 1) * 2,
            "NACQ" => (Math.Pow(Math.Exp((Math.Log(nearDiscountFactor) - Math.Log(farDiscountFactor)) / dT), 1.0 / 4.0) - 1) * 4,
            "NACM" => (Math.Pow(Math.Exp((Math.Log(nearDiscountFactor) - Math.Log(farDiscountFactor)) / dT), 1.0 / 12.0) - 1) * 12,
            _ => CommonUtils.DExcelErrorMessage($"Invalid compounding convention: {compoundingConvention}")
        };
    }

    /// <summary>
    /// Converts an interest rate from one compounding convention to another.
    /// </summary>
    /// <param name="rate">Interest rate.</param>
    /// <param name="rateTenor">The tenor associated with the supplied interest rate, inputted as a year fraction.</param>
    /// <param name="oldCompoundingConvention">Compounding convention to convert from.</param>
    /// <param name="newCompoundingConvention">Compounding convention to convert to.</param>
    /// <returns>Interest rate converted to new compounding convention.</returns>
    [ExcelFunction(
       Name = "d.IR_ConvertInterestRate",
       Description = "Converts an interest rate from one compounding convention to another.\n" +
                     "Deprecates AQS Function: 'IntConvert'",
       Category = "∂Excel: Interest Rates")]
    public static object ConvertInterestRate(
        [ExcelArgument(
                Name = "Rate",
                Description = "The interest rate to convert.")]
        double rate,
        [ExcelArgument(
                Name = "Rate Tenor",
                Description = "The tenor of the supplied interest rate to convert, as a fraction of a year.")]
        double rateTenor,
        [ExcelArgument(
                Name = "Old Compounding Convention ",
                Description = "The compounding convention to convert from.\n" +
                              "Options:'Simple', 'NACC', 'NACA', 'NACS', 'NACQ', 'NACM'")]
        string oldCompoundingConvention,
        [ExcelArgument(
                Name = "New Compounding Convention ",
                Description = "The compounding convention to convert from.\n" +
                              "Options: 'Simple', 'NACC', 'NACA', 'NACS', 'NACQ', 'NACM'")]
        string newCompoundingConvention)
    {
#if DEBUG
        CommonUtils.InFunctionWizard();
#endif
        switch (oldCompoundingConvention.ToUpper())
        {
            // First we do all conversions from simple rate to other.
            case "SIMPLE" when newCompoundingConvention.ToUpper() == "NACA":
                return (Math.Pow(1 + rate * rateTenor, 1.0 / rateTenor) - 1);
            case "SIMPLE" when newCompoundingConvention.ToUpper() == "NACS":
                return 2.0 * (Math.Pow(1 + rate * rateTenor, 1.0 / (2.0 * rateTenor)) - 1);
            case "SIMPLE" when newCompoundingConvention.ToUpper() == "NACQ":
                return 4.0 * (Math.Pow(1 + rate * rateTenor, 1.0 / (4.0 * rateTenor)) - 1);
            case "SIMPLE" when newCompoundingConvention.ToUpper() == "NACM":
                return 12.0 * (Math.Pow(1 + rate * rateTenor, 1.0 / (12.0 * rateTenor)) - 1);
            case "SIMPLE" when newCompoundingConvention.ToUpper() == "NACC":
                return Math.Log(1 + rate * rateTenor) / rateTenor;
            // Second we do all conversions from rates to simple.
            case "NACA" when newCompoundingConvention.ToUpper() == "SIMPLE":
                return (Math.Pow(1 + rate / 1.0, 1.0 * rateTenor) - 1) / rateTenor;
            case "NACS" when newCompoundingConvention.ToUpper() == "SIMPLE":
                return (Math.Pow(1 + rate / 2.0, 2.0 * rateTenor) - 1) / rateTenor;
            case "NACQ" when newCompoundingConvention.ToUpper() == "SIMPLE":
                return (Math.Pow(1 + rate / 4.0, 4.0 * rateTenor) - 1) / rateTenor;
            case "NACM" when newCompoundingConvention.ToUpper() == "SIMPLE":
                return (Math.Pow(1 + rate / 12.0, 12.0 * rateTenor) - 1) / rateTenor;
            case "NACC" when newCompoundingConvention.ToUpper() == "SIMPLE":
                return (Math.Exp(rate * rateTenor) - 1) / rateTenor;
            // Third we do all conversions from NACx rate to NACy (excluding NACC).
            case "NACA" when newCompoundingConvention.ToUpper() == "NACS":
                return 2.0 * (Math.Pow(1 + rate, 1.0 / 2.0) - 1);
            case "NACA" when newCompoundingConvention.ToUpper() == "NACQ":
                return 4.0 * (Math.Pow(1 + rate, 1.0 / 4.0) - 1);
            case "NACA" when newCompoundingConvention.ToUpper() == "NACM":
                return 12.0 * (Math.Pow(1 + rate, 1.0 / 12.0) - 1);
            case "NACS" when newCompoundingConvention.ToUpper() == "NACA":
                return 1.0 * (Math.Pow(1 + rate / 2.0, 2.0 / 1.0) - 1);
            case "NACS" when newCompoundingConvention.ToUpper() == "NACQ":
                return 4.0 * (Math.Pow(1 + rate / 2.0, 2.0 / 4.0) - 1);
            case "NACS" when newCompoundingConvention.ToUpper() == "NACM":
                return 12.0 * (Math.Pow(1 + rate / 2.0, 2.0 / 12.0) - 1);
            case "NACQ" when newCompoundingConvention.ToUpper() == "NACA":
                return 1.0 * (Math.Pow(1 + rate / 4.0, 4.0 / 1.0) - 1);
            case "NACQ" when newCompoundingConvention.ToUpper() == "NACS":
                return 2.0 * (Math.Pow(1 + rate / 4.0, 4.0 / 2.0) - 1);
            case "NACQ" when newCompoundingConvention.ToUpper() == "NACM":
                return 12.0 * (Math.Pow(1 + rate / 4.0, 4.0 / 12.0) - 1);
            case "NACM" when newCompoundingConvention.ToUpper() == "NACA":
                return 1.0 * (Math.Pow(1 + rate / 12.0, 12.0 / 1.0) - 1);
            case "NACM" when newCompoundingConvention.ToUpper() == "NACS":
                return 2.0 * (Math.Pow(1 + rate / 12.0, 12.0 / 2.0) - 1);
            case "NACM" when newCompoundingConvention.ToUpper() == "NACQ":
                return 4.0 * (Math.Pow(1 + rate / 12.0, 12.0 / 4.0) - 1);
            // Fourth we do all conversions from NACx rate to NACC.
            case "NACA" when newCompoundingConvention.ToUpper() == "NACC":
                return 1.0 * Math.Log(1 + rate / 1.0) ;
            case "NACS" when newCompoundingConvention.ToUpper() == "NACC":
                return 2.0 * Math.Log(1 + rate / 2.0);
            case "NACQ" when newCompoundingConvention.ToUpper() == "NACC":
                return 4.0 * Math.Log(1 + rate / 4.0);
            case "NACM" when newCompoundingConvention.ToUpper() == "NACC":
                return 12.0 * Math.Log(1 + rate / 12.0);
            // Fifth we do all conversions from NACC rate to NACx.
            case "NACC" when newCompoundingConvention.ToUpper() == "NACA":
                return 1.0 * (Math.Exp(rate / 1.0)-1);
            case "NACC" when newCompoundingConvention.ToUpper() == "NACS":
                return 2.0 * (Math.Exp(rate / 2.0) - 1);
            case "NACC" when newCompoundingConvention.ToUpper() == "NACQ":
                return 4.0 * (Math.Exp(rate / 4.0) - 1);
            case "NACC" when newCompoundingConvention.ToUpper() == "NACM":
                return 12.0 * (Math.Exp(rate / 12.0) - 1);
            // Sixth - if the same rate is used for conversion, it should output that rate
            default:
            {
                if (string.Equals(oldCompoundingConvention, newCompoundingConvention, StringComparison.CurrentCultureIgnoreCase) )
                {
                    return rate;
                }

                return CommonUtils.DExcelErrorMessage(
                    !new List<string> { "Simple", "NACC", "NACA", "NACS", "NACQ", "NACM" }.Contains(newCompoundingConvention)
                        ? $"Invalid new compounding convention: {newCompoundingConvention}"
                        : $"Invalid old compounding convention: {oldCompoundingConvention}");
            }
        }
    }
}
