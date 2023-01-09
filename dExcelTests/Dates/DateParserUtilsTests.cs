namespace dExcelTests.Dates;

using dExcel;
using dExcel.Dates;
using NUnit.Framework;
using QLNet;

public class DateParserUtilsTests
{
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
        yield return new TestCaseData("USD,ZAR").Returns(new JointCalendar(new UnitedStates(), new SouthAfrica()));
        yield return new TestCaseData("GBP,USD,ZAR")
            .Returns(new JointCalendar(new JointCalendar(new UnitedKingdom(), new UnitedStates()), new SouthAfrica()));
        yield return new TestCaseData("WRE").Returns(null);
        yield return new TestCaseData("WRE, USD").Returns(null);
        yield return new TestCaseData("USD, WRE").Returns(null);
        yield return new TestCaseData("EUR, USD, WRE").Returns(null);
        yield return new TestCaseData("WEEKENDSONLY").Returns(new WeekendsOnly());
        yield return new TestCaseData("WRE, NQP").Returns(null);
    }

    [Test]
    [TestCaseSource(nameof(CalendarTestData))]
    public Calendar? TestParseCalendar(string? calendarToParse)
    {
        return DateParserUtils.ParseCalendars(calendarToParse).calendar;
    }
}
