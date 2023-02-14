namespace dExcel.Utilities;

using System.Diagnostics.CodeAnalysis;
using ExcelDna.Integration;
using QLNet;

public static class CommonUtils
{
    /// <summary>
    /// Returns a message stating that the user is currently in the function wizard dialog in Excel rather than pre-computing values.
    /// Used when pre-computing values is expensive and the user is entering inputs one at a time.
    /// </summary>
    /// <returns>'In a function wizard.'</returns>
    public static string InFunctionWizard() => ExcelDnaUtil.IsInFunctionWizard() ? "In function wizard." : "";

    /// <summary>
    /// The prefix used by error messages from ∂Excel.
    /// </summary>
    public const string DExcelErrorPrefix = "#∂Excel Error:";

    /// <summary>
    /// Returns a ∂Excel specific error message.
    /// </summary>
    /// <param name="message">The error message.</param>
    /// <returns>∂Excel error message with ∂Excel prefix.</returns>
    public static string DExcelErrorMessage(string message) => $"{DExcelErrorPrefix} {message}";

    /// <summary>
    /// Returns an error message that the calendar is invalid. 
    /// </summary>
    /// <param name="invalidCalendar">The invalid calendar in question.</param>
    /// <returns>An error message that the calendar is invalid.</returns>
    public static string UnsupportedCalendarMessage(string invalidCalendar) =>
        DExcelErrorMessage($"Unsupported calendar: '{invalidCalendar}'");

    /// <summary>
    /// The sign, +1 for 'Call'/'C' or -1 for 'Put'/'P', if it can parse the option type.
    /// </summary>
    /// <param name="optionType">Option type: 'Call'/'C' or 'Put'/'P'.</param>
    /// <param name="sign">The sign, +1 for 'Call'/'C' or -1 for 'Put'/'P', if it can parse the option type.</param>
    /// <param name="errorMessage">The error message if it cannot parse the option type.</param>
    /// <returns>Returns TRUE if the option type is valid i.e., 'Call'/'C' or 'Put'/'P'. Otherwise FALSE.</returns>
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
                errorMessage = DExcelErrorMessage($"Invalid option type: '{optionType}'");
                return false;
        }
    }

    /// <summary>
    /// Returns the sign for a given direction (Long/Short i.e., Buy/Sell).
    /// </summary>
    /// <param name="direction">Direction: 'Long'/'L'/'Buy'/'B' or 'Short'/'S'/'Sell'.</param>
    /// <param name="sign">The sign, +1 for 'Long'/'L'/'Buy'/'B' and -1 for 'Short'/'S'/'Sell', if it can parse the
    /// direction.</param>
    /// <param name="errorMessage">The error message if it cannot parse the direction.</param>
    /// <returns>Returns TRUE if the direction is valid i.e., 'Long'/'L'/'Buy'/'B' or 'Short'/'S'/'Sell'.</returns>
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
                errorMessage = DExcelErrorMessage($"Invalid direction: '{direction}'");
                return false;
        }
    }

    /// <summary>
    /// Tries to parse the day count convention of the input string to an out parameter.  
    /// </summary>
    /// <param name="dayCountConventionParameter">The input string to parse.</param>
    /// <param name="dayCountConvention">The output day count convention.</param>
    /// <param name="errorMessage">The error message (if any).</param>
    /// <returns>True it can parse the string to a day count convention, else false.</returns>
    public static bool TryParseDayCountConvention(
        string dayCountConventionParameter, 
        [NotNullWhen(true)]out DayCounter? dayCountConvention,
        [NotNullWhen(false)]out string? errorMessage)
    {
        dayCountConvention =
            dayCountConventionParameter.ToUpper() switch
            {
                "ACT360" or "ACTUAL360" => new Actual360(),
                "ACT365" or "ACTUAL365" => new Actual365Fixed(),
                "ACTACT" or "ACTUALACTUAL" => new ActualActual(),
                "BUSINESS252" => new Business252(),
                "30360" or "THIRTY360" => new Thirty360(Thirty360.Thirty360Convention.BondBasis, null),
                _ => null,
            };

        if (dayCountConvention == null)
        {
            errorMessage = DExcelErrorMessage($"Invalid DayCountConvention: '{dayCountConventionParameter}'");
            return false;
        }
       
        errorMessage = null;
        return true;
    }
}
