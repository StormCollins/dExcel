namespace dExcel.Utilities;

using ExcelDna.Integration;

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
    /// Returns the sign for a given option type.
    /// </summary>
    /// <param name="direction">Option type: 'Call'/'C' or 'Put'/'P'.</param>
    /// <returns>A tuple with the sign as the first item where the sign is +1 for 'Call'/'C' or -1 for 'Put'/'P'.
    /// The second item in the tuple is the error message which is null if the sign is not null.</returns>
    public static (int? sign, string? errorMessage) GetSignOfOptionType(string direction)
    {
        switch (direction.ToUpper())
        {
            case "C":
            case "CALL":
                return (1, null);
            case "P":
            case "PUT":
                return (-1, null);
            default:
                return (null, DExcelErrorMessage($"Invalid option type: '{direction}'"));
        }
    }
    
    /// <summary>
    /// Returns the sign for a given direction (Long/Short i.e., Buy/Sell).
    /// </summary>
    /// <param name="direction">Direction: 'Long'/'L'/'Buy'/'B' or 'Short'/'S'/'Sell'.</param>
    /// <returns>A tuple with the sign as the first item where the sign is +1 for 'Long'/'L'/'Buy'/'B' and -1 for
    /// 'Short'/'S'/'Sell'.
    /// The second item in the tuple is the error message which is null if the sign is not null.</returns>
    public static (int? sign, string? errorMessage) GetSignOfDirection(string direction)
    {
        switch (direction.ToUpper())
        {
            case "B":
            case "BUY":
            case "L":
            case "LONG":
                return (1, null);
            case "S":
            case "SELL":
            case "SHORT":    
                return (-1, null);
            default:
                return (null, DExcelErrorMessage($"Invalid direction: '{direction}'"));
        }
    }
}
