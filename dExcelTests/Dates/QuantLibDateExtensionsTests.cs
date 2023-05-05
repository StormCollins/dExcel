using dExcel.Dates;
using NUnit.Framework;
using QL = QuantLib;

namespace dExcelTests.Dates;

[TestFixture]
public class DateExtensionsTests
{
    [TestCase(1, QL.Month.January)]
    [TestCase(2, QL.Month.February)]
    [TestCase(3, QL.Month.March)]
    [TestCase(4, QL.Month.April)]
    [TestCase(5, QL.Month.May)]
    [TestCase(6, QL.Month.June)]
    [TestCase(7, QL.Month.July)]
    [TestCase(8, QL.Month.August)]
    [TestCase(9, QL.Month.September)]
    [TestCase(10, QL.Month.October)]
    [TestCase(11, QL.Month.November)]
    [TestCase(12, QL.Month.December)]
    public void ToIntTest(int expected, QL.Month month)
    {
        Assert.AreEqual(expected, month.ToInt());     
    }
    
    [Test]
    public void ToQuantLibMonthTest()
    {
        DateTime dateTime = new(2023, 03, 31);
        Assert.AreEqual(QL.Month.March, dateTime.Month.ToQuantLibMonth());
    }
    
    [Test]
    public void ToQuantLibDateTest()
    {
        DateTime dateTime = new(2023, 03, 31);
        QL.Date date = dateTime.ToQuantLibDate();
        Assert.AreEqual(new QL.Date(31, QL.Month.March, 2023), date);
    }

    [Test]
    public void ToQuantLibDateFromOaDateTest()
    {
        double dateTime = new DateTime(2023, 03, 31).ToOADate();
        QL.Date date = new(31, QL.Month.March, 2023);
        Assert.AreEqual(date, dateTime.ToQuantLibDate());
    }

    [Test]
    public void ToQuantLibDateFromDateTimeObjectTest()
    {
        object dateTime = new DateTime(2023, 03, 31);
        QL.Date date = new(31, QL.Month.March, 2023);
        Assert.AreEqual(date, dateTime.ToQuantLibDate());
    }
    
    [Test]
    public void ToQuantLibDateFromOaObjectTest()
    {
        object dateTime = new DateTime(2023, 03, 31).ToOADate();
        QL.Date date = new(31, QL.Month.March, 2023);
        Assert.AreEqual(date, dateTime.ToQuantLibDate());
    }

    [Test]
    public void ToOaDateTest()
    {
        DateTime dateTime = new(2023, 03, 31);
        QL.Date date = dateTime.ToQuantLibDate();
        Assert.AreEqual(dateTime.ToOADate(), date.ToOaDate());
    }
}
