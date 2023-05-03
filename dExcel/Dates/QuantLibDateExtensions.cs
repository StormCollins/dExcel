﻿using QL = QuantLib;

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

    public static QL.Weekday ToQuantLibWeekday(this DayOfWeek dayOfWeek)
    {
        return dayOfWeek switch
        {
            DayOfWeek.Monday => QL.Weekday.Monday,
            DayOfWeek.Tuesday => QL.Weekday.Tuesday,
            DayOfWeek.Wednesday => QL.Weekday.Wednesday,
            DayOfWeek.Thursday => QL.Weekday.Thursday,
            DayOfWeek.Friday => QL.Weekday.Friday,
            DayOfWeek.Saturday => QL.Weekday.Saturday,
            DayOfWeek.Sunday => QL.Weekday.Sunday,
        };
    }

    /// <summary>
    /// Converts a QuantLib month to an integer e.g., Jan => 1, Feb => 2 etc.
    /// </summary>
    /// <param name="month">The QuantLib month to convert.</param>
    /// <returns>An integer between 1 and 12 representing a month.</returns>
    public static int ToInt(this QL.Month month)
    {
        return month switch
        {
            QL.Month.January => 1,
            QL.Month.February => 2,
            QL.Month.March => 3,
            QL.Month.April => 4,
            QL.Month.May => 5,
            QL.Month.June => 6,
            QL.Month.July => 7,
            QL.Month.August => 8,
            QL.Month.September => 9,
            QL.Month.October => 10,
            QL.Month.November => 11,
            QL.Month.December => 12,
        };
    }
    
    /// <summary>
    /// Converts a QuantLib date to a standard C# DateTime.
    /// </summary>
    /// <param name="date">The QuantLib date to convert.</param>
    /// <returns>A standard C# DateTime.</returns>
    public static DateTime ToDateTime(this QL.Date date)
    {
        return new DateTime(date.year(), date.month().ToInt(), date.dayOfMonth());
    }
}