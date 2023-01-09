namespace dExcelTests.Dates;

using dExcel;
using dExcel.Dates;
using dExcel.Utilities;
using NUnit.Framework;
using QLNet;

[TestFixture]
public sealed class DateUtilsTests
{
    public static IEnumerable<TestCaseData> FolDayHolidaysTestCaseData()
    {
        yield return new TestCaseData(new DateTime(2022, 01, 01))
            .Returns(new DateTime(2022, 01, 04));
        yield return new TestCaseData(new DateTime(2022, 01, 03))
            .Returns(new DateTime(2022, 01, 04));
        yield return new TestCaseData(new DateTime(2022, 01, 04))
            .Returns(new DateTime(2022, 01, 04));
    }
    
    [Test]
    [TestCaseSource(nameof(FolDayHolidaysTestCaseData))]
    public object FolDayUsingHolidaysTest(DateTime unadjustedDate)
    {
        object[,] holidays = { { "Holidays" }, { new DateTime(2022, 01, 03).ToOADate() } };
        return DateUtils.FolDay(unadjustedDate, holidays);
    }

    public static IEnumerable<TestCaseData> FolDayCalendarsTestData()
    {
        yield return new TestCaseData(new DateTime(2022, 06, 16), new object[,] {{"ZAR"}})
            .Returns(new DateTime(2022, 06, 17));
        yield return new TestCaseData(new DateTime(2022, 06, 17), new object[,] {{"WRE"}})
            .Returns(CommonUtils.UnsupportedCalendarMessage("WRE"));
    }
    
    [Test]
    [TestCaseSource(nameof(FolDayCalendarsTestData))]
    public object FolDayUsingCalendarsTest(DateTime unadjustedDate, object[,] calendars)
    {
        return DateUtils.FolDay(unadjustedDate, calendars);
    }

    public static IEnumerable<TestCaseData> ModFolDayHolidaysTestCaseData()
    {
        yield return new TestCaseData(new DateTime(2022, 01, 01))
            .Returns(new DateTime(2022, 01, 04));
        yield return new TestCaseData(new DateTime(2022, 01, 03))
            .Returns(new DateTime(2022, 01, 04));
        yield return new TestCaseData(new DateTime(2022, 01, 04))
            .Returns(new DateTime(2022, 01, 04));
        yield return new TestCaseData(new DateTime(2022, 04, 30))
            .Returns(new DateTime(2022, 04, 29));
    }
    
    [Test]
    [TestCaseSource(nameof(ModFolDayHolidaysTestCaseData))]
    public object ModFolDayTest(DateTime unadjusted)
    {
        object[,] holidays = { { "Holidays" }, { new DateTime(2022, 01, 03).ToOADate() } };
        return DateUtils.ModFolDay(unadjusted, holidays);
    }

    public static IEnumerable<TestCaseData> ModFolDayCalendarsTestData()
    {
        yield return new TestCaseData(new DateTime(2022, 12, 31), new object[,] {{"ZAR"}})
            .Returns(new DateTime(2022, 12, 30));
        yield return new TestCaseData(new DateTime(2022, 12, 31), new object[,] {{"WRE"}})
            .Returns(CommonUtils.UnsupportedCalendarMessage("WRE"));
    }

    [Test]
    [TestCaseSource(nameof(ModFolDayCalendarsTestData))]
    public object ModFolDayUsingCalendarsTest(DateTime unadjustedDate, object[,] calendars)
    {
        return DateUtils.ModFolDay(unadjustedDate, calendars);
    }
    
    [Test]
    public void ModFolDayInvalidHolidaysTest()
    {
        object[,] holidays = { {"Holidays"}, { "Invalid" } };
        Assert.Throws<ArgumentException>(
            () => DateUtils.ModFolDay(new DateTime(2022, 01, 01), holidays),
            $"{CommonUtils.DExcelErrorPrefix} Invalid date: 'Invalid'");
    }
    
    public static IEnumerable<TestCaseData> PrevDayHolidaysTestCaseData()
    {
        yield return new TestCaseData(new DateTime(2022, 01, 01))
            .Returns(new DateTime(2021, 12, 31));
        yield return new TestCaseData(new DateTime(2022, 01, 03))
            .Returns(new DateTime(2021, 12, 31));
        yield return new TestCaseData(new DateTime(2022, 01, 04))
            .Returns(new DateTime(2022, 01, 04));
    }

    [Test]
    [TestCaseSource(nameof(PrevDayHolidaysTestCaseData))]
    public object PrevDayTest(DateTime unadjusted)
    {
        object[,] holidays = { { "Holidays" }, { new DateTime(2022, 01, 03).ToOADate() } };
        return DateUtils.PrevDay(unadjusted, holidays);
    }

    public static IEnumerable<TestCaseData> PrevDayCalendarsTestCaseData()
    {
        yield return new TestCaseData(new DateTime(2022, 06, 16), new object[,] {{"ZAR"}})
            .Returns(new DateTime(2022, 06, 15));
        yield return new TestCaseData(new DateTime(2022, 06, 16), new object[,] {{"WRE"}})
            .Returns(CommonUtils.UnsupportedCalendarMessage("WRE"));
    }

    [Test]
    [TestCaseSource(nameof(PrevDayCalendarsTestCaseData))]
    public object PrevDayUsingCalendarsTest(DateTime unadjustedDate, object[,] calendars)
    {
        return DateUtils.PrevDay(unadjustedDate, calendars);     
    }
    
    public static IEnumerable<TestCaseData> AddTenorToDateTestCaseData()
    {
        yield return new TestCaseData(new DateTime(2022, 01, 04), "3m", "ZAR", "MODFOL")
            .Returns(new DateTime(2022, 04, 04));
    }

    [Test]
    [TestCaseSource(nameof(AddTenorToDateTestCaseData))]
    public object AddTenorToDateTest(DateTime date, string tenor, string? userCalendar, string userBusinessDayConvention)
    {
        return DateUtils.AddTenorToDate(date, tenor, userCalendar, userBusinessDayConvention);
    }

    [Test]
    public void TestAddTenorToDateWithInvalidCalendar()
    {
        object actual = DateUtils.AddTenorToDate(new DateTime(2022, 01, 01), "3m", "Invalid", "ModFol");
        string expected = CommonUtils.UnsupportedCalendarMessage("Invalid");
        Assert.AreEqual(expected, actual);
    }
    
    [Test]
    public void TestAddTenorToDateWithInvalidBusinessDayConvention()
    {
        object actual = DateUtils.AddTenorToDate(new DateTime(2022, 01, 01), "3m", "ZAR", "Invalid");
        const string expected = $"{CommonUtils.DExcelErrorPrefix} Unsupported business day convention: 'Invalid'";
        Assert.AreEqual(expected, actual);
    }
    
    [Test]
    public void TestGetAvailableDayCountConvention()
    {
        object[,] expected =
        {
            {"Variant 1", "Variant 2"},
            {"Fol", "Following"},
            {"ModFol", "ModifiedFollowing"},
            {"ModPrec", "ModPreceding"},
            {"Prec", "Preceding"},
        };
        
        Assert.AreEqual(expected, DateUtils.GetAvailableBusinessDayConventions());
    }
    
    public static IEnumerable<TestCaseData> BusinessDayConventionTestData()
    {
        yield return new TestCaseData("FOL").Returns(BusinessDayConvention.Following);
        yield return new TestCaseData("FOLLOWING").Returns(BusinessDayConvention.Following);
        yield return new TestCaseData("MODFOL").Returns(BusinessDayConvention.ModifiedFollowing);
        yield return new TestCaseData("MODIFIEDFOLLOWING").Returns(BusinessDayConvention.ModifiedFollowing);
        yield return new TestCaseData("MODPREC").Returns(BusinessDayConvention.ModifiedPreceding);
        yield return new TestCaseData("MODIFIEDPRECEDING").Returns(BusinessDayConvention.ModifiedPreceding);
        yield return new TestCaseData("PREC").Returns(BusinessDayConvention.Preceding);
        yield return new TestCaseData("PRECEDING").Returns(BusinessDayConvention.Preceding);
        yield return new TestCaseData("Invalid").Returns(null);
    } 
    
    [Test]
    [TestCaseSource(nameof(BusinessDayConventionTestData))]
    public BusinessDayConvention? TestParseBusinessDayConvention(string businessDayConventionToParse)
    {
        return DateParserUtils.ParseBusinessDayConvention(businessDayConventionToParse).businessDayConvention;
    }

    [Test]
    public void TestGetAvailableDayCountConventions()
    {
        object[,] expected =
        {
            { "Variant 1", "Variant 2", "Variant 3", "Variant 4" },
            { "Act360", "Actual360", "", "" },
            { "Act365", "Act365F", "Actual365", "Actual365F" },
            { "ActAct", "ActualActual", "", "" },
            { "Bus252", "Business252", "", "" },
        };
        
        Assert.AreEqual(expected, DateUtils.GetAvailableDayCountConventions());
    }
    
    public static IEnumerable<TestCaseData> DayCountConventionTestData()
    {
        yield return new TestCaseData("ACT360").Returns(new Actual360());
        yield return new TestCaseData("ACTUAL360").Returns(new Actual360()); 
        yield return new TestCaseData("ACT365").Returns(new Actual365Fixed());
        yield return new TestCaseData("ACT365F").Returns(new Actual365Fixed());
        yield return new TestCaseData("ACTUAL365").Returns(new Actual365Fixed());
        yield return new TestCaseData("ACTUAL365F").Returns(new Actual365Fixed());
        yield return new TestCaseData("ACTACT").Returns(new ActualActual());
        yield return new TestCaseData("ACTUALACTUAL").Returns(new ActualActual());
        yield return new TestCaseData("BUS252").Returns(new Business252());
        yield return new TestCaseData("BUSINESS252").Returns(new Business252());
        yield return new TestCaseData("Invalid").Returns(null);
    }

    [Test]
    [TestCaseSource(nameof(DayCountConventionTestData))]
    public DayCounter? TestParseDayCountConvention(string dayCountConventionToParse)
    {
        return DateUtils.ParseDayCountConvention(dayCountConventionToParse);
    }

    [Test]
    public void Act360Test()
    {
        DateTime d1 = new(2022, 01, 01);
        DateTime d2 = new(2022, 04, 01);
        Assert.AreEqual(d2.Subtract(d1).Days / 360.0, DateUtils.Act360(d1, d2)); 
    }

    [Test]
    public void Act364Test()
    {
        DateTime d1 = new(2022, 01, 01);
        DateTime d2 = new(2022, 04, 01);
        Assert.AreEqual(d2.Subtract(d1).Days / 364.0, DateUtils.Act364(d1, d2)); 
    }

    [Test]
    public void Act365Test()
    {
        DateTime d1 = new(2022, 01, 01);
        DateTime d2 = new(2022, 04, 01);
        Assert.AreEqual(d2.Subtract(d1).Days / 365.0, DateUtils.Act365(d1, d2)); 
    }

    [Test]
    public void Business252Test()
    {
        DateTime d1 = new(2022, 01, 01);
        DateTime d2 = new(2022, 04, 01);
        SouthAfrica southAfrica = new();
        Assert.AreEqual(southAfrica.businessDaysBetween(d1, d2) / 252.0, DateUtils.Business252(d1, d2, "ZAR")); 
    }

    [Test]
    public void UnsupportedCalendarBusiness252Test()
    {
        DateTime d1 = new(2022, 01, 01);
        DateTime d2 = new(2022, 04, 01);
        Assert.AreEqual(CommonUtils.UnsupportedCalendarMessage("WRE"), DateUtils.Business252(d1, d2, "WRE"));
    }
    
    public static IEnumerable<TestCaseData> WeekDaysTestData()
    {
        yield return new TestCaseData(new DateTime(2022, 01, 01)).Returns("Saturday");
        yield return new TestCaseData(new DateTime(2022, 01, 02)).Returns("Sunday");
        yield return new TestCaseData(new DateTime(2022, 01, 03)).Returns("Monday");
        yield return new TestCaseData(new DateTime(2022, 01, 04)).Returns("Tuesday");
        yield return new TestCaseData(new DateTime(2022, 01, 05)).Returns("Wednesday");
        yield return new TestCaseData(new DateTime(2022, 01, 06)).Returns("Thursday");
        yield return new TestCaseData(new DateTime(2022, 01, 07)).Returns("Friday");
    }

    [Test]
    [TestCaseSource(nameof(WeekDaysTestData))]
    public string TestParseWeekDay(DateTime date)
    {
        return DateUtils.GetWeekday(date);
    }

    public static IEnumerable<TestCaseData> IsBusinessDayTestCaseData()
    {
        yield return new TestCaseData(new DateTime(2022, 01, 01), "ZAR").Returns(false); 
        yield return new TestCaseData(new DateTime(2022, 01, 03), "ZAR").Returns(true); 
        yield return new TestCaseData(new DateTime(2022, 06, 16), "ZAR").Returns(false); 
        yield return new TestCaseData(new DateTime(2022, 06, 17), "ZAR").Returns(true); 
        yield return new TestCaseData(new DateTime(2022, 06, 17), "WRE")
            .Returns(CommonUtils.UnsupportedCalendarMessage("WRE"));
    }

    [Test]
    [TestCaseSource(nameof(IsBusinessDayTestCaseData))]
    public object IsBusinessDayTest(DateTime date, string calendars)
    {
        return DateUtils.IsBusinessDay(date, calendars); 
    }

    [Test]
    public void TestGetListOfCalendars()
    {
        object[,] expectedCalendars = DateUtils.GetListOfCalendars(); 
        object[,] actualCalendars = {
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

        Assert.AreEqual(expectedCalendars, actualCalendars);
    }

    [Test]
    public void GetListOfHolidaysTest()
    {
        object[,] expectedHolidays = DateUtils.GetListOfHolidays(new DateTime(2022, 01, 01), new DateTime(2022, 04, 01), "ZAR");
        object[,] actualHolidays = {{new DateTime(2022, 03, 21)}};
        Assert.AreEqual(expectedHolidays, actualHolidays);
    }

    [Test]
    public void GetListOfHolidaysThatAreEmptyTest()
    {
        object[,] expectedHolidays = DateUtils.GetListOfHolidays(new DateTime(2022, 01, 01), new DateTime(2022, 03, 01), "ZAR");
        object[,] actualHolidays = {{ }};
        Assert.AreEqual(expectedHolidays, actualHolidays);
    }

    [Test]
    public void GetListOfHolidaysInvalidCalendarTest()
    {
        object[,] expectedHolidays =
            DateUtils.GetListOfHolidays(new DateTime(2022, 01, 01), new DateTime(2022, 03, 01), "WRE");
        object[,] actualHolidays = {{ CommonUtils.UnsupportedCalendarMessage("WRE") }};
        Assert.AreEqual(expectedHolidays, actualHolidays);
    }

    [Test]
    public void GenerateScheduleTest()
    {
        object[,] actualSchedule = 
            DateUtils.GenerateSchedule(
                startDate: new DateTime(2022, 03, 31), 
                endDate: new DateTime(2022, 12, 31), 
                frequency: "3m",
                calendarsToParse: "ZAR", 
                businessDayConventionToParse: "MODFOL",
                ruleToParse: "Backward");

        object[,] expectedSchedule =
        {
            { new DateTime(2022, 03, 31) },
            { new DateTime(2022, 06, 30) },
            { new DateTime(2022, 09, 30) },
            { new DateTime(2022, 12, 30) },
        };

        Assert.AreEqual(expectedSchedule, actualSchedule);
    }

    [Test]
    public void GenerateScheduleUnsupportedCalendarTest()
    {
        object[,] actualSchedule =
            DateUtils.GenerateSchedule(
                startDate: new DateTime(2022, 03, 31),
                endDate: new DateTime(2022, 12, 31),
                frequency: "3m",
                calendarsToParse: "WRE",
                businessDayConventionToParse: "MODFOL",
                ruleToParse: "Backward");

        Assert.AreEqual(actualSchedule, new object[,] {{CommonUtils.UnsupportedCalendarMessage("WRE")}});
    }

    [Test]
    public void GenerateScheduleUnsupportedDayCountConventionTest()
    {
        object[,] actualSchedule =
            DateUtils.GenerateSchedule(
                startDate: new DateTime(2022, 03, 31),
                endDate: new DateTime(2022, 12, 31),
                frequency: "3m",
                calendarsToParse: "ZAR",
                businessDayConventionToParse: "SOMECONVENTION",
                ruleToParse: "Backward");

        Assert.AreEqual(
            expected: actualSchedule, 
            actual: new object[,] {{CommonUtils.DExcelErrorMessage($"Unsupported business day convention: 'SOMECONVENTION'")}});
    }

    [Test]
    public void GenerateScheduleUnsupportedRuleTest()
    {
        object[,] actualSchedule =
            DateUtils.GenerateSchedule(
                startDate: new DateTime(2022, 03, 31),
                endDate: new DateTime(2022, 12, 31),
                frequency: "3m",
                calendarsToParse: "ZAR",
                businessDayConventionToParse: "MODFOL",
                ruleToParse: "SomeRule");

        Assert.AreEqual(new object[,] {{ CommonUtils.DExcelErrorMessage("Unsupported rule specified: 'SomeRule'") }}, actualSchedule);
    }
}
