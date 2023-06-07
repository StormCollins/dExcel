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
}

