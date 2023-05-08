using System.Diagnostics.CodeAnalysis;
using QL = QuantLib;

namespace dExcel.Utilities;

/// <summary>
/// A collection of utility functions for parsing strings to commonly occuring types.
/// </summary>
public static class ParserUtils
{
    /// <summary>
    /// Returns the sign for a given direction (Long/Short i.e., Buy/Sell).
    /// </summary>
    /// <param name="direction">Direction: 'Long'/'L'/'Buy'/'B' or 'Short'/'S'/'Sell'.</param>
    /// <param name="sign">The sign, +1 for 'Long'/'L'/'Buy'/'B' and -1 for 'Short'/'S'/'Sell', if it can parse the
    /// direction.</param>
    /// <param name="errorMessage">The error message if it cannot parse the direction.</param>
    public static bool TryParseDirectionToSign(
        string direction, 
        [NotNullWhen(true)]out int? sign, 
        [NotNullWhen(false)]out string? errorMessage)
    {
        switch (direction.ToUpper())
        {
            case "B":
            case "BUY":
            case "L":
            case "LONG":
                sign = 1;
                errorMessage = null;
                return true;
            case "S":
            case "SELL":
            case "SHORT":    
                sign = -1;
                errorMessage = null;
                return true;
            default:
                sign = null;
                errorMessage = CommonUtils.DExcelErrorMessage($"Invalid direction: '{direction}'");
                return false;
        }
    }
    
    /// <summary>
    /// The sign, +1 for 'Call'/'C' or -1 for 'Put'/'P', if it can parse the option type.
    /// </summary>
    /// <param name="optionType">Option type: 'Call'/'C' or 'Put'/'P'.</param>
    /// <param name="sign">The sign, +1 for 'Call'/'C' or -1 for 'Put'/'P', if it can parse the option type.</param>
    /// <param name="errorMessage">The error message if it cannot parse the option type.</param>
    /// <returns>TRUE if the option type is valid i.e., 'Call'/'C' or 'Put'/'P', otherwise FALSE.</returns>
    public static bool TryParseOptionTypeToSign(
        string optionType, 
        [NotNullWhen(true)]out int? sign, 
        [NotNullWhen(false)]out string? errorMessage)
    {
        switch (optionType.ToUpper())
        {
            case "C":
            case "CALL":
                sign = 1;
                errorMessage = null;
                return true;
            case "P":
            case "PUT":
                sign = -1;
                errorMessage = null;
                return true;
            default:
                sign = null;
                errorMessage = CommonUtils.DExcelErrorMessage($"Invalid option type: '{optionType}'");
                return false;
        }
    }
     
    /// <summary>
    /// Tries to parse a string as a QuantLib compounding convention..
    /// </summary>
    /// <param name="compoundingConventionToParse">Compounding convention to parse.</param>
    /// <param name="compoundingConvention">The output compounding convention.</param>
    /// <param name="errorMessage">The error message (if any).</param>
    /// <returns>True if it can parse the string to a compounding convention, else false.</returns>
    public static bool TryParseQuantLibCompoundingConvention(
        string compoundingConventionToParse, 
        [NotNullWhen(true)]out (QL.Compounding compounding, QL.Frequency frequency)? compoundingConvention,
        [NotNullWhen(false)]out string? errorMessage)
    {
           compoundingConvention = 
               compoundingConventionToParse.ToUpper() switch
               {
                   "SIMPLE" => (QL.Compounding.Simple, QL.Frequency.Once),
                   "NACM" => (QL.Compounding.Compounded, QL.Frequency.Monthly),
                   "NACQ" => (QL.Compounding.Compounded, QL.Frequency.Quarterly),
                   "NACS" => (QL.Compounding.Compounded, QL.Frequency.Semiannual),
                   "NACA" => (QL.Compounding.Compounded, QL.Frequency.Annual),
                   "NACC" => (QL.Compounding.Continuous, QL.Frequency.NoFrequency),
                   _ => null,
               };

        if (compoundingConvention == null)
        {
            errorMessage = 
                CommonUtils.DExcelErrorMessage($"Invalid compounding convention: '{compoundingConventionToParse}'");
            return false;
        }
       
        errorMessage = null;
        return true;
    }
     
    /// <summary>
    /// Tries to parse the day count convention of the input string to the <see cref="dayCountConventionToParse"/> out
    /// parameter. If it cannot parse the day count convention it returns false and populates the
    /// <see cref="errorMessage"/> out parameter.
    /// </summary>
    /// <param name="dayCountConventionToParse">The input string to parse.</param>
    /// <param name="dayCountConvention">The output day count convention.</param>
    /// <param name="errorMessage">The error message (if any).</param>
    /// <returns>True it can parse the string to a day count convention, else false.</returns>
    public static bool TryParseQuantLibDayCountConvention(
        string dayCountConventionToParse, 
        [NotNullWhen(true)]out QL.DayCounter? dayCountConvention,
        [NotNullWhen(false)]out string? errorMessage)
    {
        // Note: The the 30360 convention has not been tested. 
        dayCountConvention =
            dayCountConventionToParse.ToUpper() switch
            {
                "ACT360" or "ACTUAL360" => new QL.Actual360(),
                "ACT365" or "ACTUAL365" => new QL.Actual365Fixed(),
                "ACTACT" or "ACTUALACTUAL" => new QL.ActualActual(QL.ActualActual.Convention.ISDA),
                "BUSINESS252" => new QL.Business252(),
                "30360" or "THIRTY360" => new QL.Thirty360(QL.Thirty360.Convention.ISDA),
                _ => null,
            };

        if (dayCountConvention == null)
        {
            errorMessage = CommonUtils.DExcelErrorMessage($"Invalid DayCountConvention: '{dayCountConventionToParse}'");
            return false;
        }
       
        errorMessage = null;
        return true;
    }
    
    /// <summary>
    /// Tries to parse a string to a QuantLib option type i.e, "Call" or "Put".
    /// </summary>
    /// <param name="optionType">Option type: 'Call'/'C' or 'Put'/'P'.</param>
    /// <param name="quantLibOptionType">The equivalent QuantLib option type.</param>
    /// <param name="errorMessage">The error message if it cannot parse the option type.</param>
    /// <returns>Returns TRUE if the option type is valid i.e., 'Call'/'C' or 'Put'/'P'. Otherwise FALSE.</returns>
    public static bool TryParseQuantLibOptionType(
        string optionType, 
        [NotNullWhen(true)]out QL.Option.Type? quantLibOptionType, 
        [NotNullWhen(false)]out string? errorMessage)
    {
        switch (optionType.ToUpper())
        {
            case "C":
            case "CALL":
                quantLibOptionType = QL.Option.Type.Call;
                errorMessage = null;
                return true;
            case "P":
            case "PUT":
                quantLibOptionType = QL.Option.Type.Put;
                errorMessage = null;
                return true;
            default:
                quantLibOptionType = null;
                errorMessage = CommonUtils.DExcelErrorMessage($"Invalid option type: '{optionType}'");
                return false;
        }
    }

    /// <summary>
    /// Tries to parse a string to a QuantLib swap type: 'Payer' or 'Receiver'.
    /// </summary>
    /// <param name="swapType">The string to parse.</param>
    /// <param name="qlSwapType">The output QuantLib swap type.</param>
    /// <param name="errorMessage">The error message if it cannot parse the swap type.</param>
    /// <returns>TRUE if the swap type is valid i.e., 'Payer' or 'Receiver', otherwise FALSE.</returns>
    public static bool TryParseQuantLibSwapType(
        string swapType, 
        [NotNullWhen(true)]out QL.Swap.Type? qlSwapType, 
        [NotNullWhen(false)]out string? errorMessage)
    {
        switch (swapType.ToUpper())
        {
            case "PAYER":
                qlSwapType = QL.Swap.Type.Payer;
                errorMessage = null;
                return true;
            case "RECEIVER":
                qlSwapType = QL.Swap.Type.Receiver;
                errorMessage = null;
                return true;
            default:
                qlSwapType = null;
                errorMessage = CommonUtils.DExcelErrorMessage($"Invalid swap type: '{swapType}'");
                return false;
        }
    }
}
