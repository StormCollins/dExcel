using ExcelDna.Integration;
using QL = QuantLib;
using System.Diagnostics.CodeAnalysis;

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
    /// Returns a ∂Excel specific error message if a curve parameter is missing.
    /// </summary>
    /// <param name="message">The error message.</param>
    /// <returns>∂Excel error message with ∂Excel prefix as well as the curve parameter is missing.</returns>
    public static string CurveParameterMissingErrorMessage(string missingCurveParameter) => 
        DExcelErrorMessage($"Curve parameter missing: '{missingCurveParameter}'.");
    
    /// <summary>
    /// Returns an error message that the calendar is unsupported. 
    /// </summary>
    /// <param name="unsupportedCalendar">The unsupported calendar in question.</param>
    /// <returns>An error message that the calendar is unsupported.</returns>
    public static string UnsupportedCalendarMessage(string unsupportedCalendar) =>
        DExcelErrorMessage($"Unsupported calendar: '{unsupportedCalendar}'");
}
