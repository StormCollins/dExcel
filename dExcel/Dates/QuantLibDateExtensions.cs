using QL = QuantLib;

namespace dExcel.Dates;

/// <summary>
/// A collection of extension methods for working with QuantLib dates.
/// </summary>
public static class QuantLibDateExtensions
{
    /// <summary>
    /// Converts a <see cref="DateTime"/> to a QuantLib Date.
    /// </summary>
    /// <param name="dateTime">The DateTime value to convert.</param>
    /// <returns>A QuantLib Date.</returns>
    public static QL.Date ToQuantLibDate(this DateTime dateTime)
    {
        return new QL.Date(dateTime.Day, dateTime.Month.ToQuantLibMonth(), dateTime.Year);
    }

    /// <summary>
    /// Converts an integer value, between 1 and 12, to a QuantLib Month.
    /// </summary>
    /// <param name="i">The integer value to convert.</param>
    /// <returns>A QuantLib Month.</returns>
    /// <exception cref="ArgumentOutOfRangeException">Thrown if the value is not between 1 and 12.</exception>
    public static QL.Month ToQuantLibMonth(this int i)
    {
        return i switch
        {
            1 => QL.Month.January,
            2 => QL.Month.February,
            3 => QL.Month.March,
            4 => QL.Month.April,
            5 => QL.Month.May,
            6 => QL.Month.June,
            7 => QL.Month.July,
            8 => QL.Month.August,
            9 => QL.Month.September,
            10 => QL.Month.October,
            11 => QL.Month.November,
            12 => QL.Month.December,
            _ => throw new ArgumentOutOfRangeException(nameof(i), @"Month parameter must be between 1 and 12.")
        };
    }
}
