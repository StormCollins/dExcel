﻿using dExcel.Utilities;
using ExcelDna.Integration;
using QL = QuantLib;

namespace dExcel.Dates;

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
        QL.Actual360 dayCounter = new();
        return dayCounter.yearFraction(startDate.ToQuantLibDate(), endDate.ToQuantLibDate());
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
        QL.Actual364 dayCounter = new();
        return dayCounter.yearFraction(startDate.ToQuantLibDate(), endDate.ToQuantLibDate());
    }

    /// <summary>
    /// Calculates year fraction, using the Business/252 day count convention, between two dates.
    /// </summary>
    /// <param name="startDate">Start date.</param>
    /// <param name="endDate">End date.</param>
    /// <param name="calendarsToParse">The calendars to use for the business dates.</param>
    /// <returns>The year fraction between the two dates.</returns>
    [ExcelFunction(
        Name = "d.Dates_Business252",
        Description = "Calculates year fraction, using the Business/252 day count convention, between two dates.",
        Category = "∂Excel: Dates")]
    public static object Business252(DateTime startDate, DateTime endDate, string calendarsToParse)
    {
        (QL.Calendar? calendar, string errorMessage) = DateParserUtils.ParseCalendars(calendarsToParse);
        if (calendar is null)
        {
            return errorMessage;
        }
        
        QL.Business252 dayCounter = new(calendar);
        
        return dayCounter.yearFraction(startDate.ToQuantLibDate(), endDate.ToQuantLibDate());
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
        QL.Thirty360 dayCounter = new(QL.Thirty360.Convention.ISDA);
        return dayCounter.yearFraction(startDate.ToQuantLibDate(), endDate.ToQuantLibDate());
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
        QL.Actual365Fixed dayCounter = new();
        return dayCounter.yearFraction(startDate.ToQuantLibDate(), endDate.ToQuantLibDate());
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
        (QL.Calendar? calendar, string errorMessage) result;
        if (holidaysOrCalendar.GetLength(0) == 1 && holidaysOrCalendar.GetLength(1) == 1)
        {
            result = DateParserUtils.ParseCalendars(holidaysOrCalendar[0, 0].ToString());
        }
        else
        {
            QL.BespokeCalendar bespokeCalendar = new("BespokeCalendar");
            bespokeCalendar.addWeekend(DayOfWeek.Saturday.ToQuantLibWeekday());
            bespokeCalendar.addWeekend(DayOfWeek.Sunday.ToQuantLibWeekday());
            result = DateParserUtils.ParseHolidays(holidaysOrCalendar, bespokeCalendar);
        }

        if (result.calendar is null)
        {
            return result.errorMessage;
        }

        return result.calendar?.adjust(date.ToQuantLibDate());
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
        (QL.Calendar? calendar, string errorMessage) result;
        if (holidaysOrCalendar.GetLength(0) == 1 && holidaysOrCalendar.GetLength(1) == 1)
        {
            result = DateParserUtils.ParseCalendars(holidaysOrCalendar[0, 0].ToString());
        }
        else
        {
            QL.BespokeCalendar bespokeCalendar = new("BespokeCalendar");
            bespokeCalendar.addWeekend(DayOfWeek.Saturday.ToQuantLibWeekday());
            bespokeCalendar.addWeekend(DayOfWeek.Sunday.ToQuantLibWeekday());
            result = DateParserUtils.ParseHolidays(holidaysOrCalendar, bespokeCalendar);
        }

        if (result.calendar is null)
        {
            return result.errorMessage;
        }

        return result.calendar.adjust(date.ToQuantLibDate(), QL.BusinessDayConvention.ModifiedFollowing);
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
        (QL.Calendar? calendar, string errorMessage) result;
        if (holidaysOrCalendar.GetLength(0) == 1 && holidaysOrCalendar.GetLength(1) == 1)
        {
            result = DateParserUtils.ParseCalendars(holidaysOrCalendar[0, 0].ToString());
        }
        else
        {
            QL.BespokeCalendar bespokeCalendar = new("BespokeCalendar");
            bespokeCalendar.addWeekend(DayOfWeek.Saturday.ToQuantLibWeekday());
            bespokeCalendar.addWeekend(DayOfWeek.Sunday.ToQuantLibWeekday());
            result = DateParserUtils.ParseHolidays(holidaysOrCalendar, bespokeCalendar);
        }

        if (result.calendar is null)
        {
            return result.errorMessage;
        }

        return result.calendar.adjust(date.ToQuantLibDate(), QL.BusinessDayConvention.Preceding);
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
        (QL.Calendar? calendar, string calendarErrorMessage) = DateParserUtils.ParseCalendars(userCalendar);
        if (calendar is null)
        {
            return calendarErrorMessage;
        }

        (QL.BusinessDayConvention? businessDayConvention, string errorMessage) =
            DateParserUtils.ParseBusinessDayConvention(userBusinessDayConvention);

        if (businessDayConvention is null)
        {
            return errorMessage;
        }

        if (tenor == "ON") tenor = "1d";
        if (tenor == "SW") tenor = "1w";
        
        return calendar.advance(
            d: date.ToQuantLibDate(), 
            period: new QL.Period(tenor), 
            convention: (QL.BusinessDayConvention) businessDayConvention).ToDateTime();
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
    public static QL.DayCounter? ParseDayCountConvention(string dayCountConventionToParse)
    {
        QL.DayCounter? dayCountConvention = dayCountConventionToParse.ToUpper() switch
        {
            "ACT360" or "ACTUAL360" => new QL.Actual360(),
            "ACT365" or "ACT365F" or "ACTUAL365" or "ACTUAL365F" => new QL.Actual365Fixed(),
            "ACTACT" or "ACTUALACTUAL" => new QL.ActualActual(QL.ActualActual.Convention.ISDA),
            "BUS252" or "BUSINESS252" => new QL.Business252(),
            _ => null,
        };

        return dayCountConvention;
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
            {"WEEKENDSONLY", "", "", ""},
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
        (QL.Calendar? calendar, string errorMessage) = DateParserUtils.ParseCalendars(calendarsToParse);
        if (calendar is null)
        {
            return new object[,] {{errorMessage}};
        }

        List<DateTime> holidays = new();

        for (int i = 0; i <= endDate.Subtract(startDate).Days; i++)
        {
            DateTime currentDate = startDate.AddDays(i);

            if (!calendar.isWeekend(currentDate.DayOfWeek.ToQuantLibWeekday()) && 
                calendar.isHoliday(currentDate.ToQuantLibDate()))
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
        (QL.Calendar? calendar, string calendarErrorMessage) = DateParserUtils.ParseCalendars(calendarsToParse);
        if (calendar is null)
        {
            return new object[,] {{calendarErrorMessage}};
        }

        (QL.BusinessDayConvention? businessDayConvention, string errorMessage) =
            DateParserUtils.ParseBusinessDayConvention(businessDayConventionToParse);

        if (businessDayConvention is null)
        {
            return new object[,] {{errorMessage}};
        }

        if (ruleToParse.ToUpper()!= "BACKWARD" && ruleToParse.ToUpper() != "FORWARD" && ruleToParse.ToUpper() != "IMM") 
        {
            return new object[,] {{ CommonUtils.DExcelErrorMessage($"Unsupported rule specified: '{ruleToParse}'") }};
        }

        QL.DateGeneration.Rule rule = ruleToParse.ToUpper() switch
        {
            "BACKWARD" => QL.DateGeneration.Rule.Backward,
            "FORWARD" => QL.DateGeneration.Rule.Forward,
            "IMM" => QL.DateGeneration.Rule.TwentiethIMM,
            _ => QL.DateGeneration.Rule.Forward,
        };

        QL.Schedule schedule =
            new(
                effectiveDate: startDate.ToQuantLibDate(),
                terminationDate: endDate.ToQuantLibDate(),
                tenor: new QL.Period(frequency),
                calendar: calendar,
                convention: (QL.BusinessDayConvention) businessDayConvention,
                terminationDateConvention: (QL.BusinessDayConvention) businessDayConvention,
                rule: rule,
                endOfMonth: false);

        object[,] output = new object[schedule.dates().Count, 1];
        for (int i = 0; i < schedule.dates().Count; i++)
        {
            output[i, 0] = schedule.dates()[i];
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
    
    /// <summary>
    /// Checks if the given date, for the given calendar, is a business day.
    /// </summary>
    /// <param name="date">Date to check.</param>
    /// <param name="calendarsToParse">The calendar(s).</param>
    /// <returns>True, if the given date is a business day, otherwise false.</returns>
    [ExcelFunction(
        Name = "d.Dates_IsBusinessDay",
        Description = "Checks if the given date, for the given calendar, is a business day.",
        Category = "∂Excel: Dates")]
    public static object IsBusinessDay(DateTime date, string calendarsToParse)
    {
        (QL.Calendar? calendar, string errorMessage) = DateParserUtils.ParseCalendars(calendarsToParse); 
        if (calendar is null)
        {
            return errorMessage;
        }
       
        return calendar.isBusinessDay(date.ToQuantLibDate());
    }
} 
