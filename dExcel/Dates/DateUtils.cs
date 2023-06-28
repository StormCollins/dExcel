using dExcel.Utilities;
using ExcelDna.Integration;
using QL = QuantLib;
using System.Text.RegularExpressions;

namespace dExcel.Dates;

using System.Diagnostics.CodeAnalysis;

/// <summary>
/// A collection of date utility functions.
/// </summary>
public static class DateUtils
{
    /// <summary>
    /// In Excel the column containing holidays for manual input usually has a title to the effect of 'Holidays' or
    /// 'Dates'.
    /// </summary>
    private const string ValidHolidayTitlePattern = @"(?i)(holidays?)|(dates?)|(calendar)(?-i)";
    
    /// <summary>
    /// Parses a string to a business day convention in QLNet. Users can get available business day conventions from
    /// <see cref="DateUtils.GetAvailableBusinessDayConventions"/>.
    /// </summary>
    /// <param name="businessDayConventionToParse">Business day convention to parse.</param>
    /// <returns>QLNet business day convention.</returns>
    public static (QL.BusinessDayConvention? businessDayConvention, string errorMessage) ParseBusinessDayConvention(
        string businessDayConventionToParse)
    {
        QL.BusinessDayConvention? businessDayConvention = businessDayConventionToParse.ToUpper() switch
        {
            "FOL" or "FOLLOWING" => QL.BusinessDayConvention.Following,
            "MODFOL" or "MODIFIEDFOLLOWING" => QL.BusinessDayConvention.ModifiedFollowing,
            "MODPREC" or "MODIFIEDPRECEDING" => QL.BusinessDayConvention.ModifiedPreceding,
            "PREC" or "PRECEDING" => QL.BusinessDayConvention.Preceding,
            _ => null,
        };

        return businessDayConvention is null
            ? (null, CommonUtils.DExcelErrorMessage($"Unsupported business day convention: '{businessDayConventionToParse}'"))
            : (businessDayConvention, "");
    }

    /// <summary>
    /// Used to parse a range of Excel dates to a custom QLNet calendar.
    /// </summary>
    /// <param name="holidaysOrCalendars">Holiday range.</param>
    /// <param name="calendar">The output calendar.</param>
    /// <param name="errorMessage">The output error message.</param>
    /// <returns>True if it can parse the holidays, otherwise false.</returns>
    /// <exception cref="ArgumentException">Thrown for invalid dates in <param name="holidaysOrCalendars"></param>.
    /// </exception>
    public static bool TryParseHolidays(
        object[,] holidaysOrCalendars, 
        QL.Calendar baseCalendar,
        out QL.Calendar? calendar, 
        out string errorMessage)
    {
        // There is a single column of holidays.
        if (holidaysOrCalendars.GetLength(1) == 1)
        {
            foreach (object holiday in holidaysOrCalendars)
            {
                if (double.TryParse(holiday.ToString(), out double holidayValue))
                {
                    baseCalendar.addHoliday(DateTime.FromOADate(holidayValue).ToQuantLibDate());
                }
                else
                {
                    if (!Regex.IsMatch(holiday.ToString() ?? string.Empty, ValidHolidayTitlePattern))
                    {
                        calendar = null;
                        errorMessage = CommonUtils.DExcelErrorMessage($"Invalid date: '{holiday}'");
                        return false;
                    }
                }
            }
        }
        // There are multiple columns of holidays, each with a column header specified by a specific currency/country.
        else
        {
            for (int j = 0; j < holidaysOrCalendars.GetLength(1); j++)
            {
                for (int i = 0; i < holidaysOrCalendars.GetLength(0); i++)
                {
                    if (double.TryParse(holidaysOrCalendars[i, j].ToString(), out double holidayValue))
                    {
                        baseCalendar.addHoliday(DateTime.FromOADate(holidayValue).ToQuantLibDate());
                    }
                }
            }
        }

        errorMessage = "";
        calendar = baseCalendar;
        return true;
    }

    /// <summary>
    /// Parses a string as a QLNet calendar.
    /// </summary>
    /// <param name="calendarToParse">Calendar to parse.</param>
    /// <param name="calendar"></param>
    /// <param name="errorMessage"></param>
    /// <returns>QLNet calendar.</returns>
    private static bool TryParseSingleCalendar(
        string? calendarToParse, 
        [NotNullWhen(true)]out QL.Calendar? calendar, 
        out string errorMessage)
    {
        calendar = calendarToParse?.ToUpper() switch
        {
            "ARS" or "ARGENTINA" => new QL.Argentina(),
            "AUD" or "AUSTRALIA" => new QL.Australia(),
            "AUSTRIA" => new QL.Austria(),
            "BRL" or "BRAZIL" => new QL.Brazil(),
            "BWP" or "BOTSWANA" => new QL.Botswana(),
            "CAD" or "CANADA" => new QL.Canada(),
            "CHF" or "SWITZERLAND" => new QL.Switzerland(),
            "CLP" or "CHILE" => new QL.Chile(),
            "CNH" or "CNY" or "CHINA" => new QL.China(),
            "CZK" or "CZECH REPUBLIC" => new QL.CzechRepublic(),
            "DKK" or "DENMARK" => new QL.Denmark(),
            "EUR" or "TARGET" => new QL.TARGET(),
            "GBP" or "UK" or "UNITED KINGDOM" => new QL.UnitedKingdom(),
            "FINLAND" => new QL.Finland(),
            "FRANCE" => new QL.France(),
            "GERMANY" => new QL.Germany(),
            "HKD" or "HONG KONG" => new QL.HongKong(),
            "HUF" or "HUNGARY" => new QL.Hungary(),
            "INR" or "INDIA" => new QL.India(),
            "ILS" or "ISRAEL" => new QL.Israel(),
            "IDR" or "INDONESIA" => new QL.Indonesia(),
            "ISK" or "ICELAND" => new QL.Iceland(),
            "ITALY" => new QL.Italy(),
            "JPY" or "JAPAN" => new QL.Japan(),
            "KRW" or "SOUTH KOREA" => new QL.SouthKorea(),
            "MXN" or "MEXICO" => new QL.Mexico(),
            "NOK" or "NORWAY" => new QL.Norway(),
            "NZD" or "NEW ZEALAND" => new QL.NewZealand(),
            "PLN" or "POLAND" => new QL.Poland(),
            "RON" or "ROMANIA" => new QL.Romania(),
            "RUB" or "RUSSIA" => new QL.Russia(),
            "SAR" or "SAUDI ARABIA" => new QL.SaudiArabia(),
            "SGD" or "SINGAPORE" => new QL.Singapore(),
            "SKK" or "SWEDEN" => new QL.Sweden(),
            "SLOVAKIA" => new QL.Slovakia(),
            "THB" or "THAILAND" => new QL.Thailand(),
            "TRY" or "TURKEY" => new QL.Turkey(),
            "TWD" or "TAIWAN" => new QL.Taiwan(),
            "UAH" or "UKRAINE" => new QL.Ukraine(),
            "USD" or "USA" or "UNITED STATES" or "UNITED STATES OF AMERICA" => 
                new QL.UnitedStates(QL.UnitedStates.Market.FederalReserve),
            "WEEKENDSONLY" => new QL.WeekendsOnly(),
            "ZAR" or "SOUTH AFRICA" => new QL.SouthAfrica(),
            _ => null,
        };

        if (calendar is null)
        {
            errorMessage = CommonUtils.UnsupportedCalendarMessage(calendarToParse ?? "");
            return false;
        }
        else
        {
            errorMessage = "";
            return true;
        }
    }

    /// <summary>
    /// Parses a comma-delimited list of calendars e.g., 'EUR,USD,ZAR', and creates a joint calendar.
    /// </summary>
    /// <param name="calendarsToParse">String of comma separated calendars e.g., 'EUR,USD,ZAR'.</param>
    /// <param name="calendar"></param>
    /// <param name="errorMessage"></param>
    /// <returns>A tuple consisting of the joint calendar and a possible error message.</returns>
    private static bool TryParseJointCalendar(
        string? calendarsToParse, 
        [NotNullWhen(true)]out QL.Calendar? calendar, 
        out string errorMessage)
    {
        IEnumerable<string>? calendars = calendarsToParse?.Split(',').Select(x => x.Trim());

        if (calendars != null)
        {
            IEnumerable<string> enumerable = calendars as string[] ?? calendars.ToArray();
            // (QL.Calendar? calendar0, string errorMessage0) = ParseSingleCalendar(enumerable.ElementAt(0));
            // (QL.Calendar? calendar1, string errorMessage1) = ParseSingleCalendar(enumerable.ElementAt(1));

            if (!TryParseSingleCalendar(enumerable.ElementAt(0), out QL.Calendar? calendar0, out string errorMessage0))
            {
                calendar = calendar0;
                errorMessage = errorMessage0;
                return false;
            }

            if (!TryParseSingleCalendar(enumerable.ElementAt(1), out QL.Calendar? calendar1, out string errorMessage1))
            {
                calendar = calendar1;
                errorMessage = errorMessage1;
                return false;
            }

            QL.JointCalendar jointCalendar = new(calendar0, calendar1);

            for (int i = 2; i < enumerable.Count(); i++)
            {
                if (!TryParseSingleCalendar(
                        calendarToParse: enumerable.ElementAt(i), 
                        calendar: out QL.Calendar? currentCalendar, 
                        errorMessage: out string currentErrorMessage))
                {
                    calendar = currentCalendar;
                    errorMessage = currentErrorMessage;
                    return false;
                }

                jointCalendar = new QL.JointCalendar(jointCalendar, currentCalendar);
            }

            calendar = jointCalendar;
            errorMessage = "";
            return true;
        }

        calendar = null;
        errorMessage = CommonUtils.DExcelErrorMessage("No valid calendars found.");
        return false;
    }

    /// <summary>
    /// Parses a string containing either a single or multiple calendars e.g., 'ZAR' or 'EUR,USD,ZAR'.
    /// </summary>
    /// <param name="calendarsToParse">The calendar string to parse e.g., 'ZAR' or 'EUR,USD,ZAR'.</param>
    /// <param name="calendar">The output QuantLib calendar.</param>
    /// <param name="errorMessage">The output error message.</param>
    /// <returns>A tuple containing the relevant calendar object and a possible error message.</returns>
    public static bool TryParseCalendars(
        string? calendarsToParse, 
        [NotNullWhen(true)]out QL.Calendar? calendar, 
        out string errorMessage)
    {
        if (calendarsToParse != null && calendarsToParse.Contains(','))
        {
            return TryParseJointCalendar(calendarsToParse, out calendar, out errorMessage);
        }

        return TryParseSingleCalendar(calendarsToParse, out calendar, out errorMessage);
    }
        
    /// <summary>
    /// Calculates the year fraction, using the Actual/360 day count convention, between two dates.
    /// </summary>
    /// <param name="startDate">Start date.</param>
    /// <param name="endDate">End date.</param>
    /// <returns>The Actual/360 year fraction.</returns>
    [ExcelFunction(
        Name = "d.Dates_Act360",
        Description = "Calculates year fraction, using the Actual/360 day count convention, between two dates.",
        Category = "∂Excel: Dates")]
    public static double Act360(
        [ExcelArgument(Name = "Start Date", Description = "Start date.")]DateTime startDate,
        [ExcelArgument(Name = "End Date", Description = "End date.")]DateTime endDate)
    {
        QL.Actual360 dayCounter = new();
        return dayCounter.yearFraction(startDate.ToQuantLibDate(), endDate.ToQuantLibDate());
    }
    
    /// <summary>
    /// Calculates the year fraction, using the Actual/364 day count convention, between two dates.
    /// </summary>
    /// <param name="startDate">Start date.</param>
    /// <param name="endDate">End date.</param>
    /// <returns>The Actual/364 year fraction.</returns>
    [ExcelFunction(
        Name = "d.Dates_Act364",
        Description = "Calculates year fraction, using the Actual/365 day count convention, between two dates.",
        Category = "∂Excel: Dates")]
    public static double Act364(
        [ExcelArgument(Name = "Start Date", Description = "Start date.")]DateTime startDate,
        [ExcelArgument(Name = "End Date", Description = "End date.")]DateTime endDate)
    {
        QL.Actual364 dayCounter = new();
        return dayCounter.yearFraction(startDate.ToQuantLibDate(), endDate.ToQuantLibDate());
    }

    /// <summary>
    /// Calculates the year fraction, using the Business/252 day count convention, between two dates.
    /// </summary>
    /// <param name="startDate">Start date.</param>
    /// <param name="endDate">End date.</param>
    /// <param name="calendarsToParse">The calendars to use for the business dates.</param>
    /// <returns>The year fraction between the two dates.</returns>
    [ExcelFunction(
        Name = "d.Dates_Business252",
        Description = "Calculates year fraction, using the Business/252 day count convention, between two dates.",
        Category = "∂Excel: Dates")]
    public static object Business252(
        [ExcelArgument(Name = "Start Date", Description = "Start date.")]
        DateTime startDate,
        [ExcelArgument(Name = "End Date", Description = "End date.")]
        DateTime endDate,
        [ExcelArgument(
            Name = "Calendars", 
            Description = "A comma separated list of calendars to use for the business dates.")]
        string calendarsToParse)
    {
        if (!TryParseCalendars(calendarsToParse, out QL.Calendar? calendar, out string errorMessage))
        {
            return errorMessage;
        }
        
        QL.Business252 dayCounter = new(calendar);
        return dayCounter.yearFraction(startDate.ToQuantLibDate(), endDate.ToQuantLibDate());
    }
    
    /// <summary>
    /// Calculates the year fraction, using the 30/360 day count convention, between two dates.
    /// </summary>
    /// <param name="startDate">Start date.</param>
    /// <param name="endDate">End date.</param>
    /// <returns>The 30/360 year fraction.</returns>
    [ExcelFunction(
        Name = "d.Dates_Thirty360",
        Description = "Calculates year fraction, using the 30/360 day count convention, between two dates.",
        Category = "∂Excel: Dates")]
    public static double Thirty360(
        [ExcelArgument(Name = "Start Date", Description = "Start date.")]DateTime startDate,
        [ExcelArgument(Name = "End Date", Description = "End date.")]DateTime endDate)
    {
        QL.Thirty360 dayCounter = new(QL.Thirty360.Convention.ISDA);
        return dayCounter.yearFraction(startDate.ToQuantLibDate(), endDate.ToQuantLibDate());
    }
    
    /// <summary>
    /// Calculates the year fraction, using the Actual/365 day count convention, between two dates.
    /// </summary>
    /// <param name="startDate">Start date.</param>
    /// <param name="endDate">End date.</param>
    /// <returns>The Actual/365 year fraction.</returns>
    [ExcelFunction(
        Name = "d.Dates_Act365",
        Description = "Calculates year fraction, using the Actual/365 day count convention, between two dates.",
        Category = "∂Excel: Dates")]
    public static double Act365(
        [ExcelArgument(Name = "Start Date", Description = "Start date.")]DateTime startDate,
        [ExcelArgument(Name = "End Date", Description = "End date.")]DateTime endDate)
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
        // QL.Calendar? calendar, string errorMessage) result;
        QL.Calendar? calendar;
        string errorMessage;
        if (holidaysOrCalendar.GetLength(0) == 1 && holidaysOrCalendar.GetLength(1) == 1)
        {
            if (!TryParseCalendars(holidaysOrCalendar[0, 0].ToString(), out calendar, out errorMessage))
            {
                return errorMessage;
            }
        }
        else
        {
            QL.BespokeCalendar bespokeCalendar = new("BespokeCalendar");
            bespokeCalendar.addWeekend(DayOfWeek.Saturday.ToQuantLibWeekday());
            bespokeCalendar.addWeekend(DayOfWeek.Sunday.ToQuantLibWeekday());
            if (!TryParseHolidays(holidaysOrCalendar, bespokeCalendar, out calendar, out errorMessage))
            {
                return errorMessage;
            }
        }

        QL.Date? adjustedDate = calendar?.adjust(date.ToQuantLibDate());
        if (adjustedDate is not null)
        {
            return adjustedDate.ToDateTime().ToOADate();
        }

        return CommonUtils.DExcelErrorMessage($"Invalid date: {date}");
    }

    /// <summary>
    /// Calculates the next business day using the 'modified following' convention.
    /// </summary>
    /// <param name="date">The date to adjust.</param>
    /// <param name="holidaysOrCalendar">The list of holiday dates or a calendar string (e.g., 'USD', 'ZAR' or
    /// 'USD,ZAR').</param>
    /// <returns>Adjusted business day.</returns>
    [ExcelFunction(
        Name = "d.Dates_ModFolDay",
        Description = 
            "Calculates the next business day using the 'modified following' convention.\n" +
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
        // (QL.Calendar? calendar, string errorMessage) result;
        QL.Calendar? calendar;
        string errorMessage;
        if (holidaysOrCalendar.GetLength(0) == 1 && holidaysOrCalendar.GetLength(1) == 1)
        {
            if (!TryParseCalendars(holidaysOrCalendar[0, 0].ToString(), out calendar, out errorMessage))
            {
                return errorMessage;
            }
        }
        else
        {
            QL.BespokeCalendar bespokeCalendar = new("BespokeCalendar");
            bespokeCalendar.addWeekend(DayOfWeek.Saturday.ToQuantLibWeekday());
            bespokeCalendar.addWeekend(DayOfWeek.Sunday.ToQuantLibWeekday());
            if (!TryParseHolidays(holidaysOrCalendar, bespokeCalendar, out calendar, out errorMessage))
            {
                return errorMessage;
            }
        }

        return calendar.adjust(date.ToQuantLibDate(), QL.BusinessDayConvention.ModifiedFollowing).ToOaDate();
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
        Description = 
            "Calculates the previous business day using the 'previous' convention.\n" +
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
        // (QL.Calendar? calendar, string errorMessage) result;
        QL.Calendar? calendar;
        string errorMessage;
        if (holidaysOrCalendar.GetLength(0) == 1 && holidaysOrCalendar.GetLength(1) == 1)
        {
            if (!TryParseCalendars(holidaysOrCalendar[0, 0].ToString(), out calendar, out errorMessage))
            {
                return errorMessage;
            }
        }
        else
        {
            QL.BespokeCalendar bespokeCalendar = new("BespokeCalendar");
            bespokeCalendar.addWeekend(DayOfWeek.Saturday.ToQuantLibWeekday());
            bespokeCalendar.addWeekend(DayOfWeek.Sunday.ToQuantLibWeekday());
            if (!TryParseHolidays(holidaysOrCalendar, bespokeCalendar, out calendar, out errorMessage))
            {
                return errorMessage;
            }
        }


        return calendar.adjust(date.ToQuantLibDate(), QL.BusinessDayConvention.Preceding).ToOaDate();
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
        // (QL.Calendar? calendar, string calendarErrorMessage) = TryParseCalendars(userCalendar);
        if (!TryParseCalendars(userCalendar, out QL.Calendar? calendar, out string calendarErrorMessage))
        {
            return calendarErrorMessage;
        }

        (QL.BusinessDayConvention? businessDayConvention, string errorMessage) = 
            ParseBusinessDayConvention(userBusinessDayConvention);

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
            {"", "Austria", "", ""},
            {"BWP", "Botswana", "", ""},
            {"BRL", "Brazil", "", ""},
            {"CAD", "Canada", "", ""},
            {"CHF", "Switzerland", "", ""},
            {"CLP", "Chile", "", ""},
            {"CNH", "China", "", ""},
            {"CZK", "Czech Republic", "", ""},
            {"DKK", "Denmark", "", ""},
            {"EUR", "Euro", "", ""},
            {"", "Finland", "", ""},
            {"", "France", "", ""},
            {"GBP", "Great Britain", "", ""},
            {"", "Germany", "", ""},
            {"HUF", "Hungary", "", ""},
            {"INR", "India", "", ""},
            {"ILS", "Israel", "", ""},
            {"", "Italy", "", ""},
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
            {"", "Slovakia", "", ""},
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
        // (QL.Calendar? calendar, string errorMessage) = TryParseCalendars(calendarsToParse);
        if (!TryParseCalendars(calendarsToParse, out QL.Calendar? calendar, out string errorMessage))
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
        // (QL.Calendar? calendar, string calendarErrorMessage) = TryParseCalendars(calendarsToParse);
        if (!TryParseCalendars(calendarsToParse, out QL.Calendar? calendar, out string calendarErrorMessage))
        {
            return new object[,] {{calendarErrorMessage}};
        }

        (QL.BusinessDayConvention? businessDayConvention, string errorMessage) =
            ParseBusinessDayConvention(businessDayConventionToParse);

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
            output[i, 0] = schedule.dates()[i].ToDateTime().ToOADate();
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
        // (QL.Calendar? calendar, string errorMessage) = TryParseCalendars(calendarsToParse); 
        if (!TryParseCalendars(calendarsToParse, out QL.Calendar? calendar, out string errorMessage))
        {
            return errorMessage;
        }
       
        return calendar.isBusinessDay(date.ToQuantLibDate());
    }
} 
