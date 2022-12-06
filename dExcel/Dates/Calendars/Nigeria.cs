namespace dExcel.Dates.Calendars;

using System.Globalization;
using ExcelDna.Integration;
using NodaTime;
using NodaTime.Calendars;

public static class Nigeria
{
   [ExcelFunction(
      Name = "d.IslamicCalendar")]
   public static DateTime Test()
   {
      HijriCalendar calendars = new HijriCalendar();
      var x = CalendarSystem.GetIslamicCalendar(IslamicLeapYearPattern.Base15, IslamicEpoch.Astronomical);
      DateTime d = new(2022, 12, 05);
      var hijriMonth = calendars.GetMonth(d);
      var hijriYear = calendars.GetDayOfMonth(d);
      calendars.GetDaysInMonth(1444, 9);
      return d;
      
   }
}
