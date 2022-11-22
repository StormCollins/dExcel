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
    /// <param name="holidaysOrCalendar">The list of holiday dates or a calendar string (e.g., 'USD', 'ZAR' or
    /// 'USD,ZAR').</param>
    /// <returns>Adjusted business day.</returns>
    [ExcelFunction(
        Name = "d.Date_FolDay",
        Description = "Calculates the next business day using the 'following' convention.\n" +
                      "Deprecates AQS Function: 'FolDay'",
        Category = "∂Excel: Dates")]
    public static object FolDay(
        [ExcelArgument(Name = "Date", Description = "Date to adjust.")]
        DateTime date,
        [ExcelArgument(
            Name = "Holidays/Calendar",
            Description = "The list of holiday dates or a calendar string (e.g., 'USD', 'ZAR' or 'USD,ZAR').")]
        object[,] holidaysOrCalendar)
    {
#if DEBUG
        CommonUtils.InFunctionWizard();
#endif
        (Calendar? calendar, string errorMessage) result;
        if (holidaysOrCalendar.GetLength(0) == 1 && holidaysOrCalendar.GetLength(1) == 1)
        {
            result = ParseCalendars(holidaysOrCalendar[0, 0].ToString());
        }
        else
        {
            result = ParseHolidays(holidaysOrCalendar, new WeekendsOnly());
        }

        if (result.calendar is null)
        {
            return new object[,] {{result.errorMessage}};
        }
        
        return (DateTime)result.calendar?.adjust(date);
    }

    /// <summary>
    /// Calculates the next business day using the modified following convention.
    /// </summary>
    /// <param name="date">The date to adjust.</param>
    /// <param name="holidaysOrCalendar">The list of holiday dates or a calendar string (e.g., 'USD', 'ZAR' or
    /// 'USD,ZAR').</param>
    /// <returns>Adjusted business day.</returns>
    [ExcelFunction(
        Name = "d.Date_ModFolDay",
        Description = "Calculates the next business day using the 'modified following' convention.\n" +
                      "Deprecates AQS function: 'ModFolDay'",
        Category = "∂Excel: Dates")]
    public static object ModFolDay(
        [ExcelArgument(Name = "Date", Description = "The date to adjust.")]
        DateTime date,
        [ExcelArgument(
            Name = "Holidays/Calendar", 
            Description = "The list of holiday dates or a calendar string (e.g., 'USD', 'ZAR' or 'USD,ZAR').")]
        object[,] holidaysOrCalendar)
    {
#if DEBUG
        CommonUtils.InFunctionWizard();
#endif
        (Calendar? calendar, string errorMessage) result;
        if (holidaysOrCalendar.GetLength(0) == 1 && holidaysOrCalendar.GetLength(1) == 1)
        {
            result = ParseCalendars(holidaysOrCalendar[0, 0].ToString());
        }
        else
        {
            result = ParseHolidays(holidaysOrCalendar, new WeekendsOnly());
        }

        if (result.calendar is null)
        {
            return result.errorMessage;
        }
        
        return (DateTime)result.calendar.adjust(date, BusinessDayConvention.ModifiedFollowing);
    }
        
    /// <summary>
    /// Calculates the previous business day using the 'previous' convention.
    /// </summary>
    /// <param name="date">The date to adjust.</param>
    /// <param name="holidaysOrCalendar">The list of holiday dates or a calendar string (e.g., 'USD', 'ZAR' or
    /// 'USD,ZAR').</param>
    /// <returns>The adjusted date.</returns>
    [ExcelFunction(
        Name = "d.Date_PrevDay",
        Description = "Calculates the previous business day using the 'previous' convention.\n" +
                      "Deprecates AQS function: 'PrevDay'",
        Category = "∂Excel: Dates")]
    public static object PrevDay(
        [ExcelArgument(Name = "Date", Description = "The date to adjust.")]
        DateTime date,
        [ExcelArgument(
            Name = "Holidays/Calendar", 
            Description = "The list of holiday dates or a calendar string (e.g., 'USD', 'ZAR' or 'USD,ZAR').")]
        object[,] holidaysOrCalendar)
    {
#if DEBUG
        CommonUtils.InFunctionWizard();
#endif
        (Calendar? calendar, string errorMessage) result;
        if (holidaysOrCalendar.GetLength(0) == 1 && holidaysOrCalendar.GetLength(1) == 1)
        {
            result = ParseCalendars(holidaysOrCalendar[0, 0].ToString());
        }
        else
        {
            result = ParseHolidays(holidaysOrCalendar, new WeekendsOnly());
        }

        if (result.calendar is null)
        {
            return result.errorMessage;
        }
        
        return (DateTime)result.calendar.adjust(date, BusinessDayConvention.Preceding);
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
        string? userCalendar, 
        [ExcelArgument(Name = "BDC", Description = "Business Day Convention e.g., 'MODFOL'.")]
        string userBusinessDayConvention)
    {
#if DEBUG
        CommonUtils.InFunctionWizard();
#endif
        (Calendar? calendar, string errorMessage) = ParseSingleCalendar(userCalendar);
        if (calendar is null)
        {
            return CommonUtils.DExcelErrorMessage("Invalid/unsupported calendar '{userCalendar}'.");
        }
        
        BusinessDayConvention? businessDayConvention = ParseBusinessDayConvention(userBusinessDayConvention);
        if (businessDayConvention is null)
        {
            return
                CommonUtils.DExcelErrorMessage(
                    $"Invalid/unsupported business day convention '{userBusinessDayConvention}'.");
        }
        
        return (DateTime)calendar.advance((Date)date, new Period(tenor), (BusinessDayConvention)businessDayConvention);
    }
        
    /// <summary>
    /// Used to parse a range of Excel dates to a custom QLNet calendar.
    /// </summary>
    /// <param name="holidaysOrCalendars">Holiday range.</param>
    /// <param name="calendar">Calendar.</param>
    /// <returns>A custom QLNet calendar.</returns>
    /// <exception cref="ArgumentException">Thrown for invalid dates in <param name="holidaysOrCalendars"></param>.</exception>
    private static (Calendar? calendar, string errorMessage) ParseHolidays(object[,] holidaysOrCalendars, Calendar calendar)
    {
        foreach (object holiday in holidaysOrCalendars)
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

        return (calendar, "");
    }

    /// <summary>
    /// Returns the list of available business day conventions so that a user can view them in Excel.
    /// </summary>
    /// <returns>List of available business day conventions.</returns>
    [ExcelFunction(
        Name = "d.Date_GetAvailableBusinessDayConventions",
        Description = "Lists available business day conventions in ∂Excel.",
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
        Description = "Lists available day count conventions in ∂Excel.",
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

    /// <summary>
    /// Parses a string as a QLNet calendar.
    /// </summary>
    /// <param name="calendarToParse">Calendar to parse.</param>
    /// <returns>QLNet calendar.</returns>
    public static (Calendar? calendar, string errorMessage) ParseSingleCalendar(string? calendarToParse)
    {
        Calendar? calendar = calendarToParse?.ToUpper() switch
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

        if (calendar is not null)
        {
            return (calendar, "");
        }

        return (null, CommonUtils.DExcelErrorMessage($"Unknown calendar: {calendarToParse}"));
    }
    
    /// <summary>
    /// Get a list of holidays between, and including, two dates and excluding weekends. If a holiday falls on a weekend
    /// it will not be included in this list.
    /// </summary>
    /// <param name="startDate">Start date.</param>
    /// <param name="endDate">End date</param>
    /// <param name="calendarsToParse">Calendars to parse.</param>
    /// <returns>List of holidays.</returns>
    [ExcelFunction(
        Name = "d.Date_GetListOfHolidays", 
        Description = "Get a list of holidays between, and including, two dates and excluding weekends.\n" +
                      "Note, if a holiday falls on a weekend it will not be included in this list.",
        Category = "∂Excel: Dates")]
    public static object[,] GetListOfHolidays(DateTime startDate, DateTime endDate, string calendarsToParse)
    {
        (Calendar? calendar, string errorMessage) = ParseCalendars(calendarsToParse);
        if (calendar is null)
        {
            return new object[,] {{ errorMessage }};
        }
        
        List<DateTime> holidays = new();
        
        for (int i = 0; i <= endDate.Subtract(startDate).Days; i++)
        {
            DateTime currentDate = startDate.AddDays(i);
            
            if (!calendar.isWeekend(currentDate.DayOfWeek) && calendar.isHoliday(currentDate))
            {
                holidays.Add(currentDate);
            }
        }

        if (holidays.Count == 0)
        {
            return new object[,] {{ }};
        }

        object[,] output = new object[holidays.Count, 1];
        for (int i = 0; i < holidays.Count; i++)
        {
            output[i, 0] = holidays[i];
        }

        return output;
    }
    
    private static (Calendar? calendar, string errorMessage) ParseJointCalendar(string? calendarsToParse)
    {
        IEnumerable<string>? calendars = calendarsToParse?.Split(',').Select(x => x.Trim());

        if (calendars != null)
        {
            IEnumerable<string> enumerable = calendars as string[] ?? calendars.ToArray();
            (Calendar? calendar0, string errorMessage0) = ParseSingleCalendar(enumerable.ElementAt(0));
            (Calendar? calendar1, string errorMessage1) = ParseSingleCalendar(enumerable.ElementAt(1));

            if (calendar0 is null)
            {
                return (calendar0, errorMessage0);
            }
            
            if (calendar1 is null)
            {
                return (calendar1, errorMessage1);
            }
            
            JointCalendar jointCalendar = new(calendar0, calendar1);

            for (int i = 2; i < enumerable.Count(); i++)
            {
                (Calendar? currentCalendar, string currentErrorMessage) = ParseSingleCalendar(enumerable.ElementAt(i));
                if (currentCalendar is null)
                {
                    return (currentCalendar, currentErrorMessage);
                }
                
                jointCalendar = new JointCalendar(jointCalendar, currentCalendar);
            }

            return (jointCalendar, "");
        }

        return (null, CommonUtils.DExcelErrorMessage("No valid calendars found."));
    }

    private static (Calendar? calendar, string errorMessage) ParseCalendars(string? calendarsToParse)
    {
        if (calendarsToParse != null && calendarsToParse.Contains(','))
        {
            return ParseJointCalendar(calendarsToParse);
        }

        return ParseSingleCalendar(calendarsToParse);
    }
}
