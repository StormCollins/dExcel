using dExcel.Dates;
using dExcel.Utilities;
using NUnit.Framework;
using QL = QuantLib;

namespace dExcelTests.Dates;

[TestFixture]
public sealed class DateUtilsTests
{
    public static IEnumerable<TestCaseData> FolDayHolidaysTestCaseData()
    {
        yield return new TestCaseData(new DateTime(2022, 01, 01))
            .Returns(new DateTime(2022, 01, 04).ToQuantLibDate());
    }
    
    [Test]
    [TestCaseSource(nameof(FolDayHolidaysTestCaseData))]
    public object FolDayUsingHolidaysTest(DateTime unadjustedDate)
    {
        object[,] holidays = { { "Holidays" }, { new DateTime(2022, 01, 03).ToOADate() } };
        return DateUtils.FolDay(unadjustedDate, holidays).ToQuantLibDate();
    }

    [Test]
    public void FolDayUsingMultipleHolidayListsTest()
    {
        object[,] holidays =
        {
            { "NGN", "TZS" }, 
            { new DateTime(2023, 02, 11).ToOADate(), new DateTime(2023, 02, 12).ToOADate() },
            { new DateTime(2023, 03, 12).ToOADate(), string.Empty },
        };
        
        QL.Date actual = DateUtils.FolDay(new DateTime(2023, 02, 11), holidays).ToQuantLibDate();
        QL.Date expected = new(13, 2.ToQuantLibMonth(), 2023);
        Assert.AreEqual(expected, actual);
    }
    
    public static IEnumerable<TestCaseData> FolDayCalendarsTestData()
    {
        yield return new TestCaseData(new DateTime(2022, 06, 16), new object[,] {{"ZAR"}})
            .Returns(new DateTime(2022, 06, 17).ToOADate());
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
            .Returns(new DateTime(2022, 01, 04).ToOADate());
        yield return new TestCaseData(new DateTime(2022, 01, 03))
            .Returns(new DateTime(2022, 01, 04).ToOADate());
        yield return new TestCaseData(new DateTime(2022, 01, 04))
            .Returns(new DateTime(2022, 01, 04).ToOADate());
        yield return new TestCaseData(new DateTime(2022, 04, 30))
            .Returns(new DateTime(2022, 04, 29).ToOADate());
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
            .Returns(new DateTime(2022, 12, 30).ToOADate());
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
            .Returns(new DateTime(2021, 12, 31).ToOADate());
        yield return new TestCaseData(new DateTime(2022, 01, 03))
            .Returns(new DateTime(2021, 12, 31).ToOADate());
        yield return new TestCaseData(new DateTime(2022, 01, 04))
            .Returns(new DateTime(2022, 01, 04).ToOADate());
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
            .Returns(new DateTime(2022, 06, 15).ToOADate());
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
            .Returns(new DateTime(2022, 04, 04).ToQuantLibDate());
    }

    [Test]
    [TestCaseSource(nameof(AddTenorToDateTestCaseData))]
    public object AddTenorToDateTest(DateTime date, string tenor, string? userCalendar, string userBusinessDayConvention)
    {
        return ((DateTime)DateUtils.AddTenorToDate(date, tenor, userCalendar, userBusinessDayConvention)).ToQuantLibDate();
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
        yield return new TestCaseData("FOL").Returns(QL.BusinessDayConvention.Following);
        yield return new TestCaseData("FOLLOWING").Returns(QL.BusinessDayConvention.Following);
        yield return new TestCaseData("MODFOL").Returns(QL.BusinessDayConvention.ModifiedFollowing);
        yield return new TestCaseData("MODIFIEDFOLLOWING").Returns(QL.BusinessDayConvention.ModifiedFollowing);
        yield return new TestCaseData("MODPREC").Returns(QL.BusinessDayConvention.ModifiedPreceding);
        yield return new TestCaseData("MODIFIEDPRECEDING").Returns(QL.BusinessDayConvention.ModifiedPreceding);
        yield return new TestCaseData("PREC").Returns(QL.BusinessDayConvention.Preceding);
        yield return new TestCaseData("PRECEDING").Returns(QL.BusinessDayConvention.Preceding);
        yield return new TestCaseData("Invalid").Returns(null);
    } 
    
    [Test]
    [TestCaseSource(nameof(BusinessDayConventionTestData))]
    public QL.BusinessDayConvention? TestParseBusinessDayConvention(string businessDayConventionToParse)
    {
        return DateUtils.ParseBusinessDayConvention(businessDayConventionToParse).businessDayConvention;
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
        yield return new TestCaseData("ACT360").Returns(new QL.Actual360().name());
        yield return new TestCaseData("ACTUAL360").Returns(new QL.Actual360().name()); 
        yield return new TestCaseData("ACT365").Returns(new QL.Actual365Fixed().name());
        yield return new TestCaseData("ACT365F").Returns(new QL.Actual365Fixed().name());
        yield return new TestCaseData("ACTUAL365").Returns(new QL.Actual365Fixed().name());
        yield return new TestCaseData("ACTUAL365F").Returns(new QL.Actual365Fixed().name());
        yield return new TestCaseData("ACTACT").Returns(new QL.ActualActual(QL.ActualActual.Convention.ISDA).name());
        yield return new TestCaseData("ACTUALACTUAL").Returns(new QL.ActualActual(QL.ActualActual.Convention.ISDA).name());
        yield return new TestCaseData("BUS252").Returns(new QL.Business252().name());
        yield return new TestCaseData("BUSINESS252").Returns(new QL.Business252().name());
        yield return new TestCaseData("Invalid").Returns(null);
    }

    [Test]
    [TestCaseSource(nameof(DayCountConventionTestData))]
    public string? TestParseDayCountConvention(string dayCountConventionToParse)
    {
        return DateUtils.ParseDayCountConvention(dayCountConventionToParse)?.name();
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
        QL.SouthAfrica southAfrica = new();
        Assert.AreEqual(southAfrica.businessDaysBetween(d1.ToQuantLibDate(), d2.ToQuantLibDate()) / 252.0, DateUtils.Business252(d1, d2, "ZAR")); 
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
            { new DateTime(2022, 03, 31).ToOADate() },
            { new DateTime(2022, 06, 30).ToOADate() },
            { new DateTime(2022, 09, 30).ToOADate() },
            { new DateTime(2022, 12, 30).ToOADate() },
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
    
     public static IEnumerable<TestCaseData> CalendarTestData()
        {
            yield return new TestCaseData("ARS").Returns(new QL.Argentina().name());
            yield return new TestCaseData("Argentina").Returns(new QL.Argentina().name()); 
            yield return new TestCaseData("AUD").Returns(new QL.Australia().name()); 
            yield return new TestCaseData("Australia").Returns(new QL.Australia().name()); 
            yield return new TestCaseData("Austria").Returns(new QL.Austria().name()); 
            yield return new TestCaseData("BRL").Returns(new QL.Brazil().name());
            yield return new TestCaseData("Brazil").Returns(new QL.Brazil().name());
            yield return new TestCaseData("Botswana").Returns(new QL.Botswana().name());
            yield return new TestCaseData("CAD").Returns(new QL.Canada().name());
            yield return new TestCaseData("Canada").Returns(new QL.Canada().name());
            yield return new TestCaseData("CHF").Returns(new QL.Switzerland().name());
            yield return new TestCaseData("CLP").Returns(new QL.Chile().name());
            yield return new TestCaseData("Chile").Returns(new QL.Chile().name());
            yield return new TestCaseData("Switzerland").Returns(new QL.Switzerland().name());
            yield return new TestCaseData("CNH").Returns(new QL.China().name());
            yield return new TestCaseData("CNY").Returns(new QL.China().name());
            yield return new TestCaseData("China").Returns(new QL.China().name());
            yield return new TestCaseData("CZK").Returns(new QL.CzechRepublic().name());
            yield return new TestCaseData("Czech Republic").Returns(new QL.CzechRepublic().name());
            yield return new TestCaseData("DKK").Returns(new QL.Denmark().name());
            yield return new TestCaseData("Denmark").Returns(new QL.Denmark().name());
            yield return new TestCaseData("EUR").Returns(new QL.TARGET().name());
            yield return new TestCaseData("TARGET").Returns(new QL.TARGET().name());
            yield return new TestCaseData("GBP").Returns(new QL.UnitedKingdom().name());
            yield return new TestCaseData("UK").Returns(new QL.UnitedKingdom().name());
            yield return new TestCaseData("United Kingdom").Returns(new QL.UnitedKingdom().name());
            yield return new TestCaseData("Finland").Returns(new QL.Finland().name());
            yield return new TestCaseData("France").Returns(new QL.France().name());
            yield return new TestCaseData("Germany").Returns(new QL.Germany().name());
            yield return new TestCaseData("HKD").Returns(new QL.HongKong().name());
            yield return new TestCaseData("Hong Kong").Returns(new QL.HongKong().name());
            yield return new TestCaseData("HUF").Returns(new QL.Hungary().name());
            yield return new TestCaseData("Hungary").Returns(new QL.Hungary().name());
            yield return new TestCaseData("ISK").Returns(new QL.Iceland().name());
            yield return new TestCaseData("Iceland").Returns(new QL.Iceland().name());
            yield return new TestCaseData("INR").Returns(new QL.India().name());
            yield return new TestCaseData("India").Returns(new QL.India().name());
            yield return new TestCaseData("IDR").Returns(new QL.Indonesia().name());
            yield return new TestCaseData("Indonesia").Returns(new QL.Indonesia().name());
            yield return new TestCaseData("ILS").Returns(new QL.Israel().name());
            yield return new TestCaseData("Israel").Returns(new QL.Israel().name());
            yield return new TestCaseData("Italy").Returns(new QL.Italy().name());
            yield return new TestCaseData("JPY").Returns(new QL.Japan().name());
            yield return new TestCaseData("Japan").Returns(new QL.Japan().name());
            yield return new TestCaseData("KRW").Returns(new QL.SouthKorea().name());
            yield return new TestCaseData("South Korea").Returns(new QL.SouthKorea().name());
            yield return new TestCaseData("MXN").Returns(new QL.Mexico().name());
            yield return new TestCaseData("Mexico").Returns(new QL.Mexico().name());
            yield return new TestCaseData("NOK").Returns(new QL.Norway().name());
            yield return new TestCaseData("Norway").Returns(new QL.Norway().name());
            yield return new TestCaseData("NZD").Returns(new QL.NewZealand().name());
            yield return new TestCaseData("New Zealand").Returns(new QL.NewZealand().name());
            yield return new TestCaseData("PLN").Returns(new QL.Poland().name());
            yield return new TestCaseData("Poland").Returns(new QL.Poland().name());
            yield return new TestCaseData("Russia").Returns(new QL.Russia().name());
            yield return new TestCaseData("SAR").Returns(new QL.SaudiArabia().name());
            yield return new TestCaseData("Saudi Arabia").Returns(new QL.SaudiArabia().name());
            yield return new TestCaseData("SGD").Returns(new QL.Singapore().name());
            yield return new TestCaseData("Singapore").Returns(new QL.Singapore().name());
            yield return new TestCaseData("SKK").Returns(new QL.Sweden().name());
            yield return new TestCaseData("Sweden").Returns(new QL.Sweden().name());
            yield return new TestCaseData("Slovakia").Returns(new QL.Slovakia().name());
            yield return new TestCaseData("THB").Returns(new QL.Thailand().name());
            yield return new TestCaseData("Thailand").Returns(new QL.Thailand().name());
            yield return new TestCaseData("TRY").Returns(new QL.Turkey().name());
            yield return new TestCaseData("Turkey").Returns(new QL.Turkey().name());
            yield return new TestCaseData("TWD").Returns(new QL.Taiwan().name());
            yield return new TestCaseData("Taiwan").Returns(new QL.Taiwan().name());
            yield return new TestCaseData("UAH").Returns(new QL.Ukraine().name());
            yield return new TestCaseData("Ukraine").Returns(new QL.Ukraine().name());
            yield return new TestCaseData("USD").Returns(new QL.UnitedStates(QL.UnitedStates.Market.FederalReserve).name());
            yield return new TestCaseData("USA").Returns(new QL.UnitedStates(QL.UnitedStates.Market.FederalReserve).name());
            yield return new TestCaseData("United States").Returns(new QL.UnitedStates(QL.UnitedStates.Market.FederalReserve).name());
            yield return new TestCaseData("United States of America")
                .Returns(new QL.UnitedStates(QL.UnitedStates.Market.FederalReserve).name());
            yield return new TestCaseData("ZAR").Returns(new QL.SouthAfrica().name());
            yield return new TestCaseData("South Africa").Returns(new QL.SouthAfrica().name());
            yield return new TestCaseData("Invalid").Returns(null);
            yield return new TestCaseData("USD,ZAR")
                .Returns(
                    new QL.JointCalendar(
                        new QL.UnitedStates(QL.UnitedStates.Market.FederalReserve), 
                        new QL.SouthAfrica()).name());
            yield return new TestCaseData("GBP,USD,ZAR")
                .Returns(
                    new QL.JointCalendar(
                        new QL.JointCalendar(
                            new QL.UnitedKingdom(), 
                            new QL.UnitedStates(QL.UnitedStates.Market.FederalReserve)), new QL.SouthAfrica()).name());
            yield return new TestCaseData("WRE").Returns(null);
            yield return new TestCaseData("WRE, USD").Returns(null);
            yield return new TestCaseData("USD, WRE").Returns(null);
            yield return new TestCaseData("EUR, USD, WRE").Returns(null);
            yield return new TestCaseData("WEEKENDSONLY").Returns(new QL.WeekendsOnly().name());
            yield return new TestCaseData("WRE, NQP").Returns(null);
        }
    
        [Test]
        [TestCaseSource(nameof(CalendarTestData))]
        public string? TestParseCalendar(string? calendarToParse)
        {
            return DateUtils.ParseCalendars(calendarToParse).calendar?.name();
        }
}
