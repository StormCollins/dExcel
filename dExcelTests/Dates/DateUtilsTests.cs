namespace dExcelTests;

using dExcel;
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
        object[] holidays = { new DateTime(2022, 01, 03).ToOADate() };
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
        object[] holidays = { new DateTime(2022, 01, 03).ToOADate() };
        return DateUtils.ModFolDay(unadjusted, holidays);
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
        object[] holidays = {new DateTime(2022, 01, 03).ToOADate()};
        return DateUtils.PrevDay(unadjusted, holidays);
    }

    public static IEnumerable<TestCaseData> AddTenorToDateTestCaseData()
    {
        yield return new TestCaseData(new DateTime(2022, 01, 04), "3m", "ZAR", "MODFOL")
            .Returns(new DateTime(2022, 04, 04));
    }

    [Test]
    [TestCaseSource(nameof(AddTenorToDateTestCaseData))]
    public object TestAddTenorToDate(DateTime date, string tenor, string userCalendar, string userBusinessDayConvention)
    {
        return DateUtils.AddTenorToDate(date, tenor, userCalendar, userBusinessDayConvention);
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
    
    [Test]
    [TestCase("FOL", BusinessDayConvention.Following)]
    [TestCase("FOLLOWING", BusinessDayConvention.Following)]
    [TestCase("MODFOL", BusinessDayConvention.ModifiedFollowing)]
    [TestCase("MODIFIEDFOLLOWING", BusinessDayConvention.ModifiedFollowing)]
    [TestCase("MODIFIEDPRECEDING", BusinessDayConvention.ModifiedPreceding)]
    [TestCase("MODIFIEDPRECEDING", BusinessDayConvention.ModifiedPreceding)]
    [TestCase("PRECEDING", BusinessDayConvention.Preceding)]
    public void TestParseBusinessDayConvention(string x, BusinessDayConvention businessDayConvention)
    {
        Assert.AreEqual(DateUtils.ParseBusinessDayConvention(x), businessDayConvention);
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
}
