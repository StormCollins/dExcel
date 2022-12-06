namespace dExcel.Dates;

using ExcelDna.Integration;
using QLNet;
using System.Text.RegularExpressions;
using System;

/// <summary>
/// A collection of date utility functions.
/// </summary>
public static class DateUtils
{
    /// <summary>
    /// Calculates year fraction, using the Actual/360 day count convention, between two dates.
    /// </summary>
    /// <param name="startDate">Start date.</param>
    /// <param name="endDate">End date.</param>
    /// <returns>The Actual/360 year fraction.</returns>
    [ExcelFunction(
        Name = "d.Dates_Act360",
        Description = "Calculates year fraction, using the Actual/360 day count convention, between two dates.",
        Category = "∂Excel: Dates")]
    public static double Act360(DateTime startDate, DateTime endDate)
    {
        Actual360 dayCounter = new();
        return dayCounter.yearFraction(startDate, endDate);
    }
    
    /// <summary>
    /// Calculates year fraction, using the Actual/364 day count convention, between two dates.
    /// </summary>
    /// <param name="startDate">Start date.</param>
    /// <param name="endDate">End date.</param>
    /// <returns>The Actual/364 year fraction.</returns>
    [ExcelFunction(
        Name = "d.Dates_Act364",
        Description = "Calculates year fraction, using the Actual/365 day count convention, between two dates.",
        Category = "∂Excel: Dates")]
    public static double Act364(DateTime startDate, DateTime endDate)
    {
        Actual364 dayCounter = new();
        return dayCounter.yearFraction(startDate, endDate);
    }
    
    /// <summary>
    /// Calculates year fraction, using the Business/252 day count convention, between two dates.
    /// </summary>
    /// <param name="startDate">Start date.</param>
    /// <param name="endDate">End date.</param>
    /// <returns>The Business/252 year fraction.</returns>
    [ExcelFunction(
        Name = "d.Dates_Business252",
        Description = "Calculates year fraction, using the Business/252 day count convention, between two dates.",
        Category = "∂Excel: Dates")]
    public static double Business252(DateTime startDate, DateTime endDate)
    {
        Business252 dayCounter = new();
        return dayCounter.yearFraction(startDate, endDate);
    }
    
    /// <summary>
    /// Calculates year fraction, using the 30/360 day count convention, between two dates.
    /// </summary>
    /// <param name="startDate">Start date.</param>
    /// <param name="endDate">End date.</param>
    /// <returns>The 30/360 year fraction.</returns>
    [ExcelFunction(
        Name = "d.Dates_Thirty360",
        Description = "Calculates year fraction, using the 30/360 day count convention, between two dates.",
        Category = "∂Excel: Dates")]
    public static double Thirty360(DateTime startDate, DateTime endDate)
    {
        Thirty360 dayCounter = new(QLNet.Thirty360.Thirty360Convention.ISDA);
        return dayCounter.yearFraction(startDate, endDate);
    }
    
    /// <summary>
    /// Calculates year fraction, using the Actual/365 day count convention, between two dates.
    /// </summary>
    /// <param name="startDate">Start date.</param>
    /// <param name="endDate">End date.</param>
    /// <returns>The Actual/365 year fraction.</returns>
    [ExcelFunction(
        Name = "d.Dates_Act365",
        Description = "Calculates year fraction, using the Actual/365 day count convention, between two dates.",
        Category = "∂Excel: Dates")]
    public static double Act365(DateTime startDate, DateTime endDate)
    {
        Actual365Fixed dayCounter = new();
        return dayCounter.yearFraction(startDate, endDate);
    }
    
    /// <summary>
    /// Calculates the next business day using the 'following' convention.
    /// </summary>
    /// <param name="date">The date to adjust.</param>
    /// <param name="holidaysOrCalendar">The list of holiday dates or a calendar string (e.g., 'USD', 'ZAR' or
    /// 'USD,ZAR').</param>
    /// <returns>Adjusted business day.</returns>
    [ExcelFunction(
        Name = "d.Dates_FolDay",
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
            result = DateParserUtils.ParseHolidays(holidaysOrCalendar, new WeekendsOnly());
        }

        if (result.calendar is null)
        {
            return new object[,] {{result.errorMessage}};
        }

        return (DateTime) result.calendar?.adjust(date);
    }

    /// <summary>
    /// Calculates the next business day using the modified following convention.
    /// </summary>
    /// <param name="date">The date to adjust.</param>
    /// <param name="holidaysOrCalendar">The list of holiday dates or a calendar string (e.g., 'USD', 'ZAR' or
    /// 'USD,ZAR').</param>
    /// <returns>Adjusted business day.</returns>
    [ExcelFunction(
        Name = "d.Dates_ModFolDay",
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
            result = DateParserUtils.ParseHolidays(holidaysOrCalendar, new WeekendsOnly());
        }

        if (result.calendar is null)
        {
            return result.errorMessage;
        }

        return (DateTime) result.calendar.adjust(date, BusinessDayConvention.ModifiedFollowing);
    }

    /// <summary>
    /// Calculates the previous business day using the 'previous' convention.
    /// </summary>
    /// <param name="date">The date to adjust.</param>
    /// <param name="holidaysOrCalendar">The list of holiday dates or a calendar string (e.g., 'USD', 'ZAR' or
    /// 'USD,ZAR').</param>
    /// <returns>The adjusted date.</returns>
    [ExcelFunction(
        Name = "d.Dates_PrevDay",
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
            result = DateParserUtils.ParseHolidays(holidaysOrCalendar, new WeekendsOnly());
        }

        if (result.calendar is null)
        {
            return result.errorMessage;
        }

        return (DateTime) result.calendar.adjust(date, BusinessDayConvention.Preceding);
    }

    /// <summary>
    /// Advance or retard a date by a given tenor.
    /// </summary>
    /// <param name="date">Date.</param>
    /// <param name="tenor">Tenor.</param>
    /// <param name="userCalendar">Calendar to use. Supports single calendars "EUR", "USD", "ZAR" and joint calendars
    /// e.g., "USD,ZAR".</param>
    /// <param name="userBusinessDayConvention">Business day convention selected by user e.g., "ModifiedFollowing",
    /// "Preceding".</param>
    /// <returns>The advanced or retarded date.</returns>
    [ExcelFunction(
        Name = "d.Dates_AddTenorToDate",
        Description = "Advance or retard a date by a given tenor.",
        Category = "∂Excel: Dates")]
    public static object AddTenorToDate(
        [ExcelArgument(Name = "Date", Description = "Date to adjust.")]
        DateTime date,
        [ExcelArgument(Name = "Tenor",
            Description = "Tenor amount by which to adjust the date e.g., '1w', '2m', '3y'.")]
        string tenor,
        [ExcelArgument(Name = "Calendar", Description =
            "Calendar to use. Supports single calendars 'EUR', 'USD', 'ZAR' " +
            "and joint calendars e.g., 'USD,ZAR'.")]
        string? userCalendar,
        [ExcelArgument(Name = "Business Day Convention", Description = "Business Day Convention e.g., 'MODFOL'.")]
        string userBusinessDayConvention)
    {
#if DEBUG
        CommonUtils.InFunctionWizard();
#endif
        (Calendar? calendar, string calendarErrorMessage) = ParseCalendars(userCalendar);
        if (calendar is null)
        {
            return calendarErrorMessage;
        }

        (BusinessDayConvention? businessDayConvention, string errorMessage) =
            ParseBusinessDayConvention(userBusinessDayConvention);

        if (businessDayConvention is null)
        {
            return errorMessage;
        }

        return (DateTime) calendar.advance((Date) date, new Period(tenor),
            (BusinessDayConvention) businessDayConvention);
    }

    /// <summary>
    /// Returns the list of available business day conventions so that a user can view them in Excel.
    /// </summary>
    /// <returns>List of available business day conventions.</returns>
    [ExcelFunction(
        Name = "d.Dates_GetAvailableBusinessDayConventions",
        Description = "Lists available business day conventions in ∂Excel.",
        Category = "∂Excel: Dates")]
    public static object[,] GetAvailableBusinessDayConventions()
    {
        return new object[,]
        {
            {"Variant 1", "Variant 2"},
            {"Fol", "Following"},
            {"ModFol", "ModifiedFollowing"},
            {"ModPrec", "ModPreceding"},
            {"Prec", "Preceding"},
        };
    }

    /// <summary>
    /// Parses a string to a business day convention in QLNet.
    /// Users can get available business day conventions from <see cref="GetAvailableBusinessDayConventions"/>.
    /// </summary>
    /// <param name="businessDayConventionToParse">Business day convention to parse.</param>
    /// <returns>QLNet business day convention.</returns>
    public static (BusinessDayConvention? businessDayConvention, string errorMessage) ParseBusinessDayConvention(
        string businessDayConventionToParse)
    {
        BusinessDayConvention? businessDayConvention = businessDayConventionToParse.ToUpper() switch
        {
            "FOL" or "FOLLOWING" => BusinessDayConvention.Following,
            "MODFOL" or "MODIFIEDFOLLOWING" => BusinessDayConvention.ModifiedFollowing,
            "MODPREC" or "MODIFIEDPRECEDING" => BusinessDayConvention.ModifiedPreceding,
            "PREC" or "PRECEDING" => BusinessDayConvention.Preceding,
            _ => null,
        };

        return businessDayConvention is null
            ? (null, CommonUtils.DExcelErrorMessage($"Unknown business day convention: {businessDayConventionToParse}"))
            : (businessDayConvention, "");
    }

    /// <summary>
    /// Returns the list of available Day Count Conventions so that a user can view them in Excel.
    /// </summary>
    /// <returns>List of available day count conventions.</returns>
    [ExcelFunction(
        Name = "d.Dates_GetAvailableDayCountConventions",
        Description = "Lists available day count conventions in ∂Excel.",
        Category = "∂Excel: Dates")]
    public static object[,] GetAvailableDayCountConventions()
    {
        return new object[,]
        {
            {"Variant 1", "Variant 2", "Variant 3", "Variant 4"},
            {"Act360", "Actual360", "", ""},
            {"Act365", "Act365F", "Actual365", "Actual365F"},
            {"ActAct", "ActualActual", "", ""},
            {"Bus252", "Business252", "", ""},
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

        return calendar is not null
            ? (calendar, "")
            : (null, CommonUtils.DExcelErrorMessage($"Unknown calendar: {calendarToParse}"));
    }

    /// <summary>
    /// Gets the list of available calendars.
    /// </summary>
    /// <returns>List of available calendars.</returns>
    [ExcelFunction(
        Name = "d.Dates_GetListOfCalendars", 
        Description = "Get the list of available calendars.",
        Category = "∂Excel: Dates")]
    public static object[,] GetListOfCalendars()
    {
        object[,] calendars = {
            {"Variant 1", "Variant 2", "Variant 3", "Variant 4"},
            {"ARS", "Argentina", "", ""},
            {"AUD", "Australia", "", ""},
            {"BWP", "Botswana", "", ""},
            {"BRL", "Brazil", "", ""},
            {"CAD", "Canada", "", ""},
            {"CHF", "Switzerland", "", ""},
            {"CNH", "China", "", ""},
            {"CZK", "Czech Republic", "", ""},
            {"DKK", "Denmark", "", ""},
            {"EUR", "Euro", "", ""},
            {"GBP", "Great Britain", "", ""},
            {"Germany", "", "", ""},
            {"HUF", "Hungary", "", ""},
            {"INR", "India", "", ""},
            {"ILS", "Israel", "", ""},
            {"Italy", "", "", ""},
            {"JPY", "Japan", "", ""},
            {"KRW", "South Korea", "", ""},
            {"MXN", "Mexico", "", ""},
            {"NOK", "Norway", "", ""},
            {"NZD", "New Zealand", "", ""},
            {"PLN", "Poland", "", ""},
            {"RON", "Romania", "", ""},
            {"RUB", "Russia", "", ""},
            {"SGD", "Singapore", "", ""},
            {"SKK", "Sweden", "", ""},
            {"SLOVAKIA", "", "", ""},
            {"THB", "Thailand", "", ""},
            {"TRY", "Turkey", "", ""},
            {"TWD", "Taiwan", "", ""},
            {"UAH", "Ukraine", "", ""},
            {"USD", "USA", "United States", "United States of America"},
            {"ZAR", "South Africa", "", ""},
        };
        
        return calendars;
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
        Name = "d.Dates_GetListOfHolidays",
        Description = "Get a list of holidays between, and including, two dates and excluding weekends.\n" +
                      "Note, if a holiday falls on a weekend it will not be included in this list.",
        Category = "∂Excel: Dates")]
    public static object[,] GetListOfHolidays(
        [ExcelArgument(
            Name = "Start Date", 
            Description = "Holidays are determined between start date, inclusive, and end date, inclusive.")]
        DateTime startDate, 
        [ExcelArgument(
            Name = "End Date", 
            Description = "Holidays are determined between start date, inclusive, and end date, inclusive.")]
        DateTime endDate,
        [ExcelArgument(
            Name = "Calendar(s) to Parse",
            Description = "The single calendar (e.g., 'USD', 'ZAR') or joint calendar (e.g., 'USD,ZAR') to parse.")]
        string calendarsToParse)
    {
        (Calendar? calendar, string errorMessage) = ParseCalendars(calendarsToParse);
        if (calendar is null)
        {
            return new object[,] {{errorMessage}};
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

    /// <summary>
    /// Parses a comma-delimited list of calendars e.g., 'EUR,USD,ZAR', and creates a joint calendar.
    /// </summary>
    /// <param name="calendarsToParse">String of comma separated calendars e.g., 'EUR,USD,ZAR'.</param>
    /// <returns>A tuple consisting of the joint calendar and a possible error message.</returns>
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

    /// <summary>
    /// Parses a string containing either a single or multiple calendars e.g., 'ZAR' or 'EUR,USD,ZAR'.
    /// </summary>
    /// <param name="calendarsToParse">The calendar string to parse e.g., 'ZAR' or 'EUR,USD,ZAR'.</param>
    /// <returns>A tuple containing the relevant calendar object and a possible error message.</returns>
    private static (Calendar? calendar, string errorMessage) ParseCalendars(string? calendarsToParse)
    {
        if (calendarsToParse != null && calendarsToParse.Contains(','))
        {
            return ParseJointCalendar(calendarsToParse);
        }

        return ParseSingleCalendar(calendarsToParse);
    }

    /// <summary>
    /// Generates a date schedule. Does not currently support stub periods.
    /// </summary>
    /// <param name="startDate">Start date.</param>
    /// <param name="endDate">End date.</param>
    /// <param name="frequency">Frequency.</param>
    /// <param name="calendarsToParse">The calendar(s).</param>
    /// <param name="businessDayConventionToParse">Business day convention.</param>
    /// <param name="ruleToParse">The date generation rule. 'Backward' = Start from end date and work backwards.
    /// 'Forward' = Start from start date and work forwards. 'IMM' = IMM dates.</param>
    /// <returns></returns>
    [ExcelFunction(
        Name = "d.Dates_GenerateSchedule",
        Description = "Generates a schedule of dates.",
        Category = "∂Excel: Dates")]
    public static object[,] GenerateSchedule(
        [ExcelArgument(Name = "Start Date", Description= "Start date.")]
        DateTime startDate,
        [ExcelArgument(Name = "End Date", Description = "End date.")]
        DateTime endDate,
        [ExcelArgument(Name = "Frequency", Description = "Frequency e.g., '3m', '6m', '1y' etc.")]
        string frequency,
        [ExcelArgument(Name = "Calendar(s)", Description = "The calendar(s) to parse e.g., 'USD', 'ZAR', 'USD,ZAR' etc.")]
        string calendarsToParse,
        [ExcelArgument(Name = "Business Day Convention", Description = "Business day convention e.g., 'FOL', 'MODFOL', 'PREC' etc.")]
        string businessDayConventionToParse,
        [ExcelArgument(
            Name = "Rule", 
            Description = "The date generation rule. " +
                          "\n'Backward' = Start from end date and move backwards. " +
                          "\n'Forward' = Start from start date and move forwards. " +
                          "\n'IMM' = IMM dates.")]
        string ruleToParse)
    {
        (Calendar? calendar, string calendarErrorMessage) = ParseCalendars(calendarsToParse);
        if (calendar is null)
        {
            return new object[,] {{calendarErrorMessage}};
        }

        (BusinessDayConvention? businessDayConvention, string errorMessage) =
            ParseBusinessDayConvention(businessDayConventionToParse);

        if (businessDayConvention is null)
        {
            return new object[,] {{errorMessage}};
        }

        if (ruleToParse.ToUpper()!= "BACKWARD" && ruleToParse.ToUpper() != "FORWARD" && ruleToParse.ToUpper() != "IMM") 
        {
            return new object[,] {{ CommonUtils.DExcelErrorMessage($"Invalid rule specified: {ruleToParse}") }};
        }

        DateGeneration.Rule rule = ruleToParse.ToUpper() switch
        {
            "BACKWARD" => DateGeneration.Rule.Backward,
            "FORWARD" => DateGeneration.Rule.Forward,
            "IMM" => DateGeneration.Rule.TwentiethIMM,
            _ => DateGeneration.Rule.Forward,
        };

        Schedule schedule =
            new(
                effectiveDate: new Date(startDate),
                terminationDate: new Date(endDate),
                tenor: new Period(frequency),
                calendar: calendar,
                convention: (BusinessDayConvention) businessDayConvention,
                terminationDateConvention: (BusinessDayConvention) businessDayConvention,
                rule: rule,
                endOfMonth: false);

        object[,] output = new object[schedule.dates().Count, 1];
        for (int i = 0; i < schedule.dates().Count; i++)
        {
            output[i, 0] = schedule.dates()[i].ToDateTime();
        }

        return output;
    }

    /// <summary>
    /// Gets the day of week for the given date.
    /// </summary>
    /// <param name="date">The date.</param>
    /// <returns>The day of the week.</returns>
    [ExcelFunction(
        Name = "d.Dates_GetWeekday", 
        Description = "Gets the day of the week for the given date.",
        Category = "∂Excel: Dates")]
    public static string GetWeekday(DateTime date)
    {
        return date.DayOfWeek.ToString();
    }
} 
