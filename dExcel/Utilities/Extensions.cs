using System.Text;
using System.Text.RegularExpressions;
using Omicron;
using QL = QuantLib;

namespace dExcel.Utilities;

/// <summary>
/// A collection of useful extension methods for various types.
/// </summary>
public static class Extensions
{
    /// <summary>
    /// Succinctly wraps the string comparison method, ignoring case.
    /// </summary>
    /// <param name="s">The current string.</param>
    /// <param name="values">The set of values to be compared to.</param>
    /// <returns>True if the two strings are the same, ignoring case, otherwise False.</returns>
    public static bool IgnoreCaseEquals<T>(this string? s, params T[] values)
    {
        foreach (T value in values)
        {
            if (string.Equals(s, value?.ToString(), StringComparison.InvariantCultureIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Adds spaces before the capitals of a camel case string (assuming it has no spaces already).
    /// E.g., "CamelCase" -> "Camel Case".
    /// It also capitalizes the very first letter e.g., "camelCase" -> "Camel Case".
    /// </summary>
    /// <param name="s">The word to added spaces to.</param>
    /// <returns>A string with spaces before the capitals.</returns>
    public static string SplitCamelCase(this string s)
    {
        StringBuilder output = new(Regex.Replace(s, "([A-Z])", " $1", RegexOptions.Compiled).Trim());
        output[0] = char.ToUpper(output[0]);
        return output.ToString();
    }

    /// <summary>
    /// Converts an Omicron Tenor to a QuantLib period.
    /// </summary>
    /// <param name="tenor">Omicron tenor to convert.</param>
    /// <returns>QuantLib period.</returns>
    /// <exception cref="ArgumentOutOfRangeException">Thrown if an invalid time unit is chosen.</exception>
    public static QuantLib.Period ToQuantLibPeriod(this Tenor tenor)
    {
        QL.TimeUnit timeUnit = tenor.Unit switch
        {
            TenorUnit.Day => QL.TimeUnit.Days,
            TenorUnit.Week => QL.TimeUnit.Weeks,
            TenorUnit.Month => QL.TimeUnit.Months,
            TenorUnit.Year => QL.TimeUnit.Years,
            _ => throw new ArgumentOutOfRangeException()
        };

        return new(tenor.Amount, timeUnit);
    }

    /// <summary>
    /// Converts a QuantLib period to an Omicron Tenor .
    /// </summary>
    /// <returns>QuantLib period.</returns>
    /// <exception cref="ArgumentOutOfRangeException">Thrown if an invalid time unit is chosen.</exception>
    public static Tenor ToOmicronTenor(this QL.Period period)
    {
        TenorUnit tenorUnit = period.units() switch
        {
            QL.TimeUnit.Days => TenorUnit.Day,
            QL.TimeUnit.Weeks => TenorUnit.Week,
            QL.TimeUnit.Months => TenorUnit.Month,
            QL.TimeUnit.Years => TenorUnit.Year,
            _ => throw new ArgumentOutOfRangeException()
        };

        return new(period.length(), tenorUnit);
    }
}
