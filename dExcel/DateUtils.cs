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
        Name = "d.Date_FolDay",
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
        Name = "d.Date_ModFolDay",
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
        Name = "d.Date_PrevDay",
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

    /// <summary>
    /// Advance or retard a date by a given tenor.
    /// </summary>
    /// <param name="date">Date.</param>
    /// <param name="tenor">Tenor.</param>
    /// <param name="userCalendar">Calendar selected by user e.g., "EUR", "USD", "ZAR".</param>
    /// <param name="userBusinessDayConvention">Business day convention selected by user e.g., "ModifiedFollowing", "Preceding".</param>
    /// <returns>The advanced or retarded date.</returns>
    [ExcelFunction(
        Name = "d.Date_AddTenorToDate",
        Description = "Advance or retard a date by a given tenor.",
        Category = "∂Excel: Dates")]
    public static object AddTenorToDate(
        [ExcelArgument(Name = "Date", Description = "Date to adjust.")]
        DateTime date, 
        [ExcelArgument(Name = "Tenor", Description = "Tenor amount by which to adjust the date e.g., '1w', '2m', '3y'.")]
        string tenor, 
        [ExcelArgument(Name = "Calendar", Description = "The calendar to use e.g., 'ZAR', 'USD'.")]
        string userCalendar, 
        [ExcelArgument(Name = "BDC", Description = "Business Day Convention e.g., 'MODFOL'.")]
        string userBusinessDayConvention)
    {
#if DEBUG
        CommonUtils.InFunctionWizard();
#endif
        Calendar? calendar = null;
        if (userCalendar.ToUpper() == "ZAR")
        {
            calendar = new SouthAfrica();
        }

        BusinessDayConvention? businessDayConvention = null; 
        if (userBusinessDayConvention.ToUpper() == "MODFOL")
        {
            businessDayConvention = BusinessDayConvention.ModifiedFollowing;
        }

        return (DateTime)calendar?.advance((Date)date, new Period(tenor), (BusinessDayConvention)businessDayConvention);
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
                if (!Regex.IsMatch(holiday.ToString() ?? string.Empty, ValidHolidayTitlePattern))
                {
                    throw new ArgumentException($"Invalid date: {holiday}");
                }
            }
        }

        return calendar;
    }

    /// <summary>
    /// Returns the list of available Business Day Conventions so that a user can peruse them in Excel.
    /// </summary>
    /// <returns>List of available Business Day Conventions.</returns>
    [ExcelFunction(
        Name = "d.Date_GetAvailableBusinessDayConventions",
        Category = "∂Excel")]
    public static object[,] GetAvailableBusinessDayConventions()
    {
        return new object[,]
        {
            { "Variant 1", "Variant 2" },
            { "Fol", "Following" },
            { "ModFol", "ModifiedFollowing" },
            { "ModPrec", "ModPreceding" },
            { "Prec", "Preceding" },
        };
    }
}
