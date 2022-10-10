namespace dExcelTests;

using dExcel;
using NUnit.Framework;

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
}
