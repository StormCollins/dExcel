using dExcel.Dates;
using NUnit.Framework;
using QL = QuantLib;

namespace dExcelTests.Dates;

public class DateParserUtilsTests
{
    public static IEnumerable<TestCaseData> CalendarTestData()
    {
        yield return new TestCaseData("ARS").Returns(new QL.Argentina().name());
        yield return new TestCaseData("Argentina").Returns(new QL.Argentina().name()); 
        yield return new TestCaseData("AUD").Returns(new QL.Australia().name());
        yield return new TestCaseData("Australia").Returns(new QL.Australia().name());
        yield return new TestCaseData("BRL").Returns(new QL.Brazil().name());
        yield return new TestCaseData("Brazil").Returns(new QL.Brazil().name());
        yield return new TestCaseData("CAD").Returns(new QL.Canada().name());
        yield return new TestCaseData("Canada").Returns(new QL.Canada().name());
        yield return new TestCaseData("CHF").Returns(new QL.Switzerland().name());
        yield return new TestCaseData("Switzerland").Returns(new QL.Switzerland().name());
        yield return new TestCaseData("CNH").Returns(new QL.China().name());
        yield return new TestCaseData("CNY").Returns(new QL.China().name());
        yield return new TestCaseData("China").Returns(new QL.China().name());
        yield return new TestCaseData("CZK").Returns(new QL.CzechRepublic().name());
        yield return new TestCaseData("Czech Republic").Returns(new QL.CzechRepublic().name());
        yield return new TestCaseData("DKK").Returns(new QL.Denmark().name());
        yield return new TestCaseData("Denmark").Returns(new QL.Denmark().name());
        yield return new TestCaseData("EUR").Returns(new QL.TARGET().name());
        yield return new TestCaseData("GBP").Returns(new QL.UnitedKingdom().name());
        yield return new TestCaseData("UK").Returns(new QL.UnitedKingdom().name());
        yield return new TestCaseData("United Kingdom").Returns(new QL.UnitedKingdom().name());
        yield return new TestCaseData("Germany").Returns(new QL.Germany().name());
        yield return new TestCaseData("HKD").Returns(new QL.HongKong().name());
        yield return new TestCaseData("Hong Kong").Returns(new QL.HongKong().name());
        yield return new TestCaseData("HUF").Returns(new QL.Hungary().name());
        yield return new TestCaseData("Hungary").Returns(new QL.Hungary().name());
        yield return new TestCaseData("INR").Returns(new QL.India().name());
        yield return new TestCaseData("India").Returns(new QL.India().name());
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
        yield return new TestCaseData("RON").Returns(new QL.Romania().name());
        yield return new TestCaseData("Romania").Returns(new QL.Romania().name());
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
        return DateParserUtils.ParseCalendars(calendarToParse).calendar?.name();
    }
}
