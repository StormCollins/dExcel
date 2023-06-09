using ExcelDna.Integration;
using QL = QuantLib;

namespace dExcel.Utilities;

/// <summary>
/// A collection of common utility functions that don't quite fit elsewhere.
/// </summary>
public static class CommonUtils
{
    /// <summary>
    /// Returns a message stating that the user is currently in the function wizard dialog in Excel rather than
    /// pre-computing values. Used when pre-computing values is expensive and the user is entering inputs one at a time.
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
    /// Returns a ∂Excel specific error message that a curve parameter is missing.
    /// </summary>
    /// <param name="missingCurveParameter">The actual name of the missing curve parameter.</param>
    /// <returns>A ∂Excel error message that a curve parameter is missing.</returns>
    /// <remarks>If, for example, a variable/parameter named "baseDate" is missing one would invoke this function with
    /// nameof(baseDate).CurveParameterMissingErrorMessage(). This would return something to the effect of
    /// "Curve parameter 'Base Date' missing." i.e., it formats it nicely and uses reflection to find the name of the
    /// parameter.</remarks>
    public static string CurveParameterMissingErrorMessage(
        this string missingCurveParameter) => 
        DExcelErrorMessage($"Missing curve parameter: '{missingCurveParameter.SplitCamelCase()}'.");
    
    /// <summary>
    /// Returns an error message that the calendar is unsupported. 
    /// </summary>
    /// <param name="unsupportedCalendar">The unsupported calendar in question.</param>
    /// <returns>An error message that the calendar is unsupported.</returns>
    public static string UnsupportedCalendarMessage(string unsupportedCalendar) =>
        DExcelErrorMessage($"Unsupported calendar: '{unsupportedCalendar}'");
}
