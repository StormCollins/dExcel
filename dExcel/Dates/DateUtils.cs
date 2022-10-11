namespace dExcel.Dates;

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
#if DEBUG
        CommonUtils.InFunctionWizard();
#endif
        Calendar calendar = ParseHolidays(holidays, new WeekendsOnly());
        return (DateTime)calendar.adjust(date);
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
        Calendar calendar = ParseHolidays(holidays, new WeekendsOnly());
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
        Calendar calendar = ParseHolidays(holidays, new WeekendsOnly());
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
        Calendar? calendar = ParseCalendar(userCalendar);
        if (calendar is null)
        {
            return $"{CommonUtils.DExcelErrorPrefix} Invalid/unsupported calendar '{userCalendar}'.";
        }
        
        BusinessDayConvention? businessDayConvention = ParseBusinessDayConvention(userBusinessDayConvention);
        if (businessDayConvention is null)
        {
            return
                $"{CommonUtils.DExcelErrorPrefix} Invalid/unsupported business day convention " +
                $"'{userBusinessDayConvention}'.";
        }
        
        return (DateTime)calendar.advance((Date)date, new Period(tenor), (BusinessDayConvention)businessDayConvention);
    }
        
    /// <summary>
    /// Used to parse a range of Excel dates to a custom QLNet calendar.
    /// </summary>
    /// <param name="holidays">Holiday range.</param>
    /// <param name="calendar">Calendar.</param>
    /// <returns>A custom QLNet calendar.</returns>
    /// <exception cref="ArgumentException">Thrown for invalid dates in <param name="holidays"></param>.</exception>
    private static Calendar ParseHolidays(object[] holidays, Calendar calendar)
    {
        foreach (object holiday in holidays)
        {
            if (double.TryParse(holiday.ToString(), out double holidayValue))
            {
                calendar.addHoliday(DateTime.FromOADate(holidayValue));
            }
            else
            {
                if (!Regex.IsMatch(holiday.ToString() ?? string.Empty, ValidHolidayTitlePattern))
                {
                    throw new ArgumentException($"{CommonUtils.DExcelErrorPrefix} Invalid date '{holiday}'.");
                }
            }
        }

        return calendar;
    }

    /// <summary>
    /// Returns the list of available Business Day Conventions so that a user can view them in Excel.
    /// </summary>
    /// <returns>List of available business day conventions.</returns>
    [ExcelFunction(
        Name = "d.Date_GetAvailableBusinessDayConventions",
        Category = "∂Excel: Dates")]
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
    
    /// <summary>
    /// Parses a string to a business day convention in QLNet.
    /// Users can get available business day conventions from <see cref="GetAvailableBusinessDayConventions"/>.
    /// </summary>
    /// <param name="businessDayConventionToParse">Business day convention to parse.</param>
    /// <returns>QLNet business day convention.</returns>
    public static BusinessDayConvention? ParseBusinessDayConvention(string businessDayConventionToParse)
    {
        BusinessDayConvention? businessDayConvention = businessDayConventionToParse.ToUpper() switch
        {
            "FOL" or "FOLLOWING" => BusinessDayConvention.Following,
            "MODFOL" or "MODIFIEDFOLLOWING" => BusinessDayConvention.ModifiedFollowing,
            "MODPREC" or "MODIFIEDPRECEDING" => BusinessDayConvention.ModifiedPreceding,
            "PREC" or "PRECEDING" => BusinessDayConvention.Preceding,
            _ => null,
        };

        return businessDayConvention;
    }

    /// <summary>
    /// Returns the list of available Day Count Conventions so that a user can view them in Excel.
    /// </summary>
    /// <returns>List of available day count conventions.</returns>
    [ExcelFunction(
        Name = "d.Date_GetAvailableDayCountConventions",
        Category = "∂Excel: Dates")]
    public static object[,] GetAvailableDayCountConventions()
    {
        return new object[,]
        {
            { "Variant 1", "Variant 2", "Variant 3", "Variant 4" },
            { "Act360", "Actual360", "", "" },
            { "Act365", "Act365F", "Actual365", "Actual365F" },
            { "ActAct", "ActualActual", "", "" },
            { "Bus252", "Business252", "", "" },
        };
    }
    
    /// <summary>
    /// Parses a string to a QLNet day counter convention.
    /// </summary>
    /// <param name="dayCountConventionToParse">Day count convention to parse.</param>
    /// <returns>QLNet day count convention.</returns>
    public static DayCounter? ParseDayCountConvention(string dayCountConventionToParse)
    {
        DayCounter? dayCountConvention = dayCountConventionToParse.ToUpper() switch
        {
            "ACT360" or "ACTUAL360" => new Actual360(),
            "ACT365" or "ACT365F" or "ACTUAL365" or "ACTUAL365F" => new Actual365Fixed(),
            "ACTACT" or "ACTUALACTUAL" => new ActualActual(),
            "BUS252" or "BUSINESS252" => new Business252(),
            _ => null,
        };

        return dayCountConvention;
    }

    public static Calendar? ParseCalendar(string calendarToParse)
    {
        Calendar? calendar = calendarToParse.ToUpper() switch
        {
            "ARS" or "ARGENTINA" => new Argentina(),
            "AUD" or "AUSTRALIA" => new Australia(),
            "BWP" or "BOTSWANA" => new Botswana(),
            "BRL" or "BRAZIL" => new Brazil(),
            "CAD" or "CANADA" => new Canada(),
            "CHF" or "SWITZERLAND" => new Switzerland(),
            "CNH" or "CNY" or "CHINA" => new China(),
            "CZK" or "CZECH REPUBLIC" => new CzechRepublic(),
            "DKK" or "DENMARK" => new Denmark(),
            "EUR" => new TARGET(),
            "GBP" or "UK" or "UNITED KINGDOM" => new UnitedKingdom(),
            "GERMANY" => new Germany(),
            "HKD" or "HONG KONG" => new HongKong(),
            "HUF" or "HUNGARY" => new Hungary(),
            "INR" or "INDIA" => new India(),
            "ILS" or "ISRAEL" => new Israel(),
            "ITALY" => new Italy(),
            "JPY" or "JAPAN" => new Japan(), 
            "KRW" or "SOUTH KOREA" => new SouthKorea(),
            "MXN" or "MEXICO" => new Mexico(),
            "NOK" or "NORWAY" => new Norway(),
            "NZD" or "NEW ZEALAND" => new NewZealand(),
            "PLN" or "POLAND" => new Poland(),
            "RON" or "ROMANIA" => new Romania(),
            "RUB" or "RUSSIA" => new Russia(),
            "SAR" or "SAUDI ARABIA" => new SaudiArabia(),
            "SGD" or "SINGAPORE" => new Singapore(),
            "SKK" or "SWEDEN" => new Sweden(),
            "SLOVAKIA" => new Slovakia(),
            "THB" or "THAILAND" => new Thailand(),
            "TRY" or "TURKEY" => new Turkey(),
            "TWD" or "TAIWAN" => new Taiwan(),
            "UAH" or "UKRAINE" => new Ukraine(),
            "USD" or "USA" or "UNITED STATES" or "UNITED STATES OF AMERICA" => new UnitedStates(),
            "ZAR" or "SOUTH AFRICA" => new SouthAfrica(),
            _ => null,
        };

        return calendar;
    }
}
