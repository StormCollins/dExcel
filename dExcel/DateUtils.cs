namespace dExcel;

using System;
using System.Text.RegularExpressions;
using ExcelDna.Integration;
using QLNet;

/// <summary>
/// A collection of date utility functions.
/// </summary>
public static class DateUtils
{
    private const string ValidHolidayTitlePattern = @"(?i)(holidays?)|(dates?)|(calendar)(?-i)";

    /// <summary>
    /// Calculates the next business day using the 'following' convention.
    /// </summary>
    /// <param name="date">The date to adjust.</param>
    /// <param name="holidays">The list of holiday dates.</param>
    /// <returns>Adjusted business day.</returns>
    [ExcelFunction(
        Name = "d.FolDay",
        Description = "Calculates the next business day using the 'following' convention.\n" +
                        "Deprecates AQS Function: 'FolDay'",
        Category = "∂Excel: Dates")]
    public static object FolDay(
        [ExcelArgument(Name = "Date", Description = "Date to adjust.")]
        DateTime date,
        [ExcelArgument(Name = "Holidays", Description = "List of holiday dates.")]
        object[] holidays)
    {
        //CommonUtils.InFunctionWizard();
        var calendar = ParseHolidays(holidays, new WeekendsOnly());
        return (DateTime)calendar.adjust(date, BusinessDayConvention.Following);
    }

    /// <summary>
    /// Calculates the next business day using the modified following convention.
    /// </summary>
    /// <param name="date">The date to adjust.</param>
    /// <param name="holidays">The list of holiday dates.</param>
    /// <returns>Adjusted business day.</returns>
    [ExcelFunction(
        Name = "d.ModFolDay",
        Description = "Calculates the next business day using the 'modified following' convention.\n" +
                        "Deprecates AQS function: 'ModFolDay'",
        Category = "∂Excel: Dates")]
    public static object ModFolDay(
        [ExcelArgument(Name = "Date", Description = "The date to adjust.")]
        DateTime date,
        [ExcelArgument(Name = "Holidays", Description = "The list of holiday dates.")]
        object[] holidays)
    {
#if DEBUG
        CommonUtils.InFunctionWizard();
#endif
        var calendar = ParseHolidays(holidays, new WeekendsOnly());
        return (DateTime)calendar.adjust(date, BusinessDayConvention.ModifiedFollowing);
    }
        
    /// <summary>
    /// Calculates the previous business day using the 'previous' convention.
    /// </summary>
    /// <param name="date">The date to adjust.</param>
    /// <param name="holidays">The list of holiday dates.</param>
    /// <returns>The adjusted date.</returns>
    [ExcelFunction(
        Name = "d.PrevDay",
        Description = "Calculates the previous business day using the 'previous' convention.\n" +
                        "Deprecates AQS function: 'PrevDay'",
        Category = "∂Excel: Dates")]
    public static object PrevDay(
        [ExcelArgument(Name = "Date", Description = "The date to adjust.")]
        DateTime date,
        [ExcelArgument(Name = "Holidays", Description = "The list of holiday dates.")]
        object[] holidays)
    {
#if DEBUG
        CommonUtils.InFunctionWizard();
#endif
        var calendar = ParseHolidays(holidays, new WeekendsOnly());
        return (DateTime)calendar.adjust(date, BusinessDayConvention.Preceding);
    }

    private static Calendar ParseHolidays(object[] holidays, Calendar calendar)
    {
        foreach (var holiday in holidays)
        {
            if (double.TryParse(holiday.ToString(), out var holidayValue))
            {
                calendar.addHoliday(DateTime.FromOADate(holidayValue));
            }
            else
            {
                if (!Regex.IsMatch(holiday.ToString(), ValidHolidayTitlePattern))
                {
                    throw new ArgumentException($"Invalid date: {holiday}");
                }
            }
        }

        return calendar;
    }
}

