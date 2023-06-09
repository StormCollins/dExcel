using System.Text.RegularExpressions;

namespace dExcel.Utilities;

using System.Text;

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

}
