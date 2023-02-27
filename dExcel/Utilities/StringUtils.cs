namespace dExcel.Utilities;

using ExcelDna.Integration;
using System.Text.RegularExpressions;

/// <summary>
/// A class containing a set of utility functions for working with strings.
/// </summary>
public static class StringUtils
{
    /// <summary>
    /// Returns a string that matches the specified regular expression (regex).
    /// </summary>
    /// <param name="input">The input string to search.</param>
    /// <param name="pattern">The regex pattern.</param>
    /// <returns>The string matched by the regex.</returns>
    [ExcelFunction(
        Name = "d.StringUtils_RegexMatch",
        Description = "Returns a string that matches the specified regular expression (regex).",
        Category = "∂Excel: String Utils")]
    public static string RegexMatch(string input, string pattern)
    {
        return Regex.Match(input, pattern, RegexOptions.IgnoreCase).Value;
    }
}
