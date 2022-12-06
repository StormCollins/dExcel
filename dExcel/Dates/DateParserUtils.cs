namespace dExcel.Dates;

using QLNet;
using System.Text.RegularExpressions;


public static class DateParserUtils
{
    /// <summary>
    /// In Excel the column containing holidays for manual input usually has a title to the effect of 'Holidays' or
    /// 'Dates'.
    /// </summary>
    private const string ValidHolidayTitlePattern = @"(?i)(holidays?)|(dates?)|(calendar)(?-i)";
    
    /// <summary>
    /// Used to parse a range of Excel dates to a custom QLNet calendar.
    /// </summary>
    /// <param name="holidaysOrCalendars">Holiday range.</param>
    /// <param name="calendar">Calendar.</param>
    /// <returns>A custom QLNet calendar.</returns>
    /// <exception cref="ArgumentException">Thrown for invalid dates in <param name="holidaysOrCalendars"></param>.
    /// </exception>
    public static (Calendar? calendar, string errorMessage) ParseHolidays(
        object[,] holidaysOrCalendars, 
        Calendar calendar)
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
}
