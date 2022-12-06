namespace dExcelTests.Dates;

using dExcel;
using dExcel.Dates;
using NUnit.Framework;
using QLNet;

[TestFixture]
public sealed class DateUtilsTests
{
    public static IEnumerable<TestCaseData> FolDayTestCaseData()
    {
        yield return new TestCaseData(new DateTime(2022, 01, 01))
            .Returns(new DateTime(2022, 01, 04));
        yield return new TestCaseData(new DateTime(2022, 01, 03))
            .Returns(new DateTime(2022, 01, 04));
        yield return new TestCaseData(new DateTime(2022, 01, 04))
            .Returns(new DateTime(2022, 01, 04));
    }
    
    [Test]
    [TestCaseSource(nameof(FolDayTestCaseData))]
    public object FolDayTest(DateTime unadjusted)
    {
        object[,] holidays = { { "Holidays" }, { new DateTime(2022, 01, 03).ToOADate() } };
        return DateUtils.FolDay(unadjusted, holidays);
    }
    
    public static IEnumerable<TestCaseData> ModFolDayTestCaseData()
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
    [TestCaseSource(nameof(ModFolDayTestCaseData))]
    public object ModFolDayTest(DateTime unadjusted)
    {
        object[,] holidays = { { "Holidays" }, { new DateTime(2022, 01, 03).ToOADate() } };
        return DateUtils.ModFolDay(unadjusted, holidays);
    }
    
    [Test]
    public void ModFolDayInvalidHolidaysTest()
    {
        object[,] holidays = { {"Holidays"}, { "Invalid" } };
        Assert.Throws<ArgumentException>(
            () => DateUtils.ModFolDay(new DateTime(2022, 01, 01), holidays),
            $"{CommonUtils.DExcelErrorPrefix} Invalid date: 'Invalid'");
    }
    
    public static IEnumerable<TestCaseData> PrevDayTestCaseData()
    {
        yield return new TestCaseData(new DateTime(2022, 01, 01))
            .Returns(new DateTime(2021, 12, 31));
        yield return new TestCaseData(new DateTime(2022, 01, 03))
            .Returns(new DateTime(2021, 12, 31));
        yield return new TestCaseData(new DateTime(2022, 01, 04))
            .Returns(new DateTime(2022, 01, 04));
    }

    [Test]
    [TestCaseSource(nameof(PrevDayTestCaseData))]
    public object PrevDayTest(DateTime unadjusted)
    {
        object[,] holidays = { { "Holidays" }, { new DateTime(2022, 01, 03).ToOADate() } };
        return DateUtils.PrevDay(unadjusted, holidays);
    }

    public static IEnumerable<TestCaseData> AddTenorToDateTestCaseData()
    {
        yield return new TestCaseData(new DateTime(2022, 01, 04), "3m", "ZAR", "MODFOL")
            .Returns(new DateTime(2022, 04, 04));
    }

    [Test]
    [TestCaseSource(nameof(AddTenorToDateTestCaseData))]
    public object TestAddTenorToDate(DateTime date, string tenor, string? userCalendar, string userBusinessDayConvention)
    {
        return DateUtils.AddTenorToDate(date, tenor, userCalendar, userBusinessDayConvention);
    }

    [Test]
    public void TestAddTenorToDateWithInvalidCalendar()
    {
        object actual = DateUtils.AddTenorToDate(new DateTime(2022, 01, 01), "3m", "Invalid", "ModFol");
        const string expected = $"{CommonUtils.DExcelErrorPrefix} Unsupported calendar: 'Invalid'";
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
    
    public static IEnumerable<TestCaseData> CalendarTestData()
    {
        yield return new TestCaseData("ARS").Returns(new Argentina());
        yield return new TestCaseData("Argentina").Returns(new Argentina()); 
        yield return new TestCaseData("AUD").Returns(new Australia());
        yield return new TestCaseData("Australia").Returns(new Australia());
        yield return new TestCaseData("BWP").Returns(new Botswana());
        yield return new TestCaseData("Botswana").Returns(new Botswana());
        yield return new TestCaseData("BRL").Returns(new Brazil());
        yield return new TestCaseData("Brazil").Returns(new Brazil());
        yield return new TestCaseData("CAD").Returns(new Canada());
        yield return new TestCaseData("Canada").Returns(new Canada());
        yield return new TestCaseData("CHF").Returns(new Switzerland());
        yield return new TestCaseData("Switzerland").Returns(new Switzerland());
        yield return new TestCaseData("CNH").Returns(new China());
        yield return new TestCaseData("CNY").Returns(new China());
        yield return new TestCaseData("China").Returns(new China());
        yield return new TestCaseData("CZK").Returns(new CzechRepublic());
        yield return new TestCaseData("Czech Republic").Returns(new CzechRepublic());
        yield return new TestCaseData("DKK").Returns(new Denmark());
        yield return new TestCaseData("Denmark").Returns(new Denmark());
        yield return new TestCaseData("EUR").Returns(new TARGET());
        yield return new TestCaseData("GBP").Returns(new UnitedKingdom());
        yield return new TestCaseData("UK").Returns(new UnitedKingdom());
        yield return new TestCaseData("United Kingdom").Returns(new UnitedKingdom());
        yield return new TestCaseData("Germany").Returns(new Germany());
        yield return new TestCaseData("HKD").Returns(new HongKong());
        yield return new TestCaseData("Hong Kong").Returns(new HongKong());
        yield return new TestCaseData("HUF").Returns(new Hungary());
        yield return new TestCaseData("Hungary").Returns(new Hungary());
        yield return new TestCaseData("INR").Returns(new India());
        yield return new TestCaseData("India").Returns(new India());
        yield return new TestCaseData("ILS").Returns(new Israel());
        yield return new TestCaseData("Israel").Returns(new Israel());
        yield return new TestCaseData("Italy").Returns(new Italy());
        yield return new TestCaseData("JPY").Returns(new Japan());
        yield return new TestCaseData("Japan").Returns(new Japan());
        yield return new TestCaseData("KRW").Returns(new SouthKorea());
        yield return new TestCaseData("South Korea").Returns(new SouthKorea());
        yield return new TestCaseData("MXN").Returns(new Mexico());
        yield return new TestCaseData("Mexico").Returns(new Mexico());
        yield return new TestCaseData("NOK").Returns(new Norway());
        yield return new TestCaseData("Norway").Returns(new Norway());
        yield return new TestCaseData("NZD").Returns(new NewZealand());
        yield return new TestCaseData("New Zealand").Returns(new NewZealand());
        yield return new TestCaseData("PLN").Returns(new Poland());
        yield return new TestCaseData("Poland").Returns(new Poland());
        yield return new TestCaseData("RON").Returns(new Romania());
        yield return new TestCaseData("Romania").Returns(new Romania());
        yield return new TestCaseData("Russia").Returns(new Russia());
        yield return new TestCaseData("SAR").Returns(new SaudiArabia());
        yield return new TestCaseData("Saudi Arabia").Returns(new SaudiArabia());
        yield return new TestCaseData("SGD").Returns(new Singapore());
        yield return new TestCaseData("Singapore").Returns(new Singapore());
        yield return new TestCaseData("SKK").Returns(new Sweden());
        yield return new TestCaseData("Sweden").Returns(new Sweden());
        yield return new TestCaseData("Slovakia").Returns(new Slovakia());
        yield return new TestCaseData("THB").Returns(new Thailand());
        yield return new TestCaseData("Thailand").Returns(new Thailand());
        yield return new TestCaseData("TRY").Returns(new Turkey());
        yield return new TestCaseData("Turkey").Returns(new Turkey());
        yield return new TestCaseData("TWD").Returns(new Taiwan());
        yield return new TestCaseData("Taiwan").Returns(new Taiwan());
        yield return new TestCaseData("UAH").Returns(new Ukraine());
        yield return new TestCaseData("Ukraine").Returns(new Ukraine());
        yield return new TestCaseData("USD").Returns(new UnitedStates());
        yield return new TestCaseData("USA").Returns(new UnitedStates());
        yield return new TestCaseData("United States").Returns(new UnitedStates());
        yield return new TestCaseData("United States of America").Returns(new UnitedStates());
        yield return new TestCaseData("ZAR").Returns(new SouthAfrica());
        yield return new TestCaseData("South Africa").Returns(new SouthAfrica());
        yield return new TestCaseData("Invalid").Returns(null);
    }

    [Test]
    [TestCaseSource(nameof(CalendarTestData))]
    public Calendar? TestParseCalendar(string? calendarToParse)
    {
        return DateParserUtils.ParseCalendars(calendarToParse).calendar;
    }
}
