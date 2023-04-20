namespace dExcel.Dates;

using QL = QuantLib;
using System.Text.RegularExpressions;
using Utilities;

public static class DateParserUtils
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
    /// <param name="calendar">Calendar.</param>
    /// <returns>A custom QLNet calendar.</returns>
    /// <exception cref="ArgumentException">Thrown for invalid dates in <param name="holidaysOrCalendars"></param>.
    /// </exception>
    public static (QL.Calendar? calendar, string errorMessage) ParseHolidays(
        object[,] holidaysOrCalendars, 
        QL.Calendar calendar)
    {
        // There is a single column of holidays.
        if (holidaysOrCalendars.GetLength(1) == 1)
        {
            foreach (object holiday in holidaysOrCalendars)
            {
                if (double.TryParse(holiday.ToString(), out double holidayValue))
                {
                    calendar.addHoliday(DateTime.FromOADate(holidayValue).ToQuantLibDate());
                }
                else
                {
                    if (!Regex.IsMatch(holiday.ToString() ?? string.Empty, ValidHolidayTitlePattern))
                    {
                        throw new ArgumentException(CommonUtils.DExcelErrorMessage($"Invalid date: '{holiday}'"));
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
                        calendar.addHoliday(DateTime.FromOADate(holidayValue).ToQuantLibDate());
                    }
                }
            }
        }

        return (calendar, "");
    }

    /// <summary>
    /// Parses a string as a QLNet calendar.
    /// </summary>
    /// <param name="calendarToParse">Calendar to parse.</param>
    /// <returns>QLNet calendar.</returns>
    private static (QL.Calendar? calendar, string errorMessage) ParseSingleCalendar(string? calendarToParse)
    {
        QL.Calendar? calendar = calendarToParse?.ToUpper() switch
        {
            "ARS" or "ARGENTINA" => new QL.Argentina(),
            "AUD" or "AUSTRALIA" => new QL.Australia(),
            "BRL" or "BRAZIL" => new QL.Brazil(),
            "CAD" or "CANADA" => new QL.Canada(),
            "CHF" or "SWITZERLAND" => new QL.Switzerland(),
            "CNH" or "CNY" or "CHINA" => new QL.China(),
            "CZK" or "CZECH REPUBLIC" => new QL.CzechRepublic(),
            "DKK" or "DENMARK" => new QL.Denmark(),
            "EUR" => new QL.TARGET(),
            "GBP" or "UK" or "UNITED KINGDOM" => new QL.UnitedKingdom(),
            "GERMANY" => new QL.Germany(),
            "HKD" or "HONG KONG" => new QL.HongKong(),
            "HUF" or "HUNGARY" => new QL.Hungary(),
            "INR" or "INDIA" => new QL.India(),
            "ILS" or "ISRAEL" => new QL.Israel(),
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

        return calendar is not null
            ? (calendar, "")
            : (null, CommonUtils.UnsupportedCalendarMessage(calendarToParse ?? ""));
    }

    /// <summary>
    /// Parses a comma-delimited list of calendars e.g., 'EUR,USD,ZAR', and creates a joint calendar.
    /// </summary>
    /// <param name="calendarsToParse">String of comma separated calendars e.g., 'EUR,USD,ZAR'.</param>
    /// <returns>A tuple consisting of the joint calendar and a possible error message.</returns>
    private static (QL.Calendar? calendar, string errorMessage) ParseJointCalendar(string? calendarsToParse)
    {
        IEnumerable<string>? calendars = calendarsToParse?.Split(',').Select(x => x.Trim());

        if (calendars != null)
        {
            IEnumerable<string> enumerable = calendars as string[] ?? calendars.ToArray();
            (QL.Calendar? calendar0, string errorMessage0) = ParseSingleCalendar(enumerable.ElementAt(0));
            (QL.Calendar? calendar1, string errorMessage1) = ParseSingleCalendar(enumerable.ElementAt(1));

            if (calendar0 is null)
            {
                return (calendar0, errorMessage0);
            }

            if (calendar1 is null)
            {
                return (calendar1, errorMessage1);
            }

            QL.JointCalendar jointCalendar = new(calendar0, calendar1);

            for (int i = 2; i < enumerable.Count(); i++)
            {
                (QL.Calendar? currentCalendar, string currentErrorMessage) = ParseSingleCalendar(enumerable.ElementAt(i));
                if (currentCalendar is null)
                {
                    return (currentCalendar, currentErrorMessage);
                }

                jointCalendar = new QL.JointCalendar(jointCalendar, currentCalendar);
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
    public static (QL.Calendar? calendar, string errorMessage) ParseCalendars(string? calendarsToParse)
    {
        if (calendarsToParse != null && calendarsToParse.Contains(','))
        {
            return ParseJointCalendar(calendarsToParse);
        }

        return ParseSingleCalendar(calendarsToParse);
    }
}
