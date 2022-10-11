namespace dExcelTests.Currencies;

using NUnit.Framework;
using dExcel.Currencies;
using QLNet;

[TestFixture]
public sealed class CurrencyUtilsTests
{
    public static IEnumerable<TestCaseData> CurrencyTestData()
    {
        yield return new TestCaseData("EUR").Returns(new EURCurrency());
        yield return new TestCaseData("GBP").Returns(new GBPCurrency());
        yield return new TestCaseData("USD").Returns(new USDCurrency());
        yield return new TestCaseData("ZAR").Returns(new ZARCurrency());
        yield return new TestCaseData("Invalid").Returns(null);
    }
    
    [Test]
    [TestCaseSource(nameof(CurrencyTestData))]
    public Currency? TestParseCurrency(string currencyToParse)
    {
        return CurrencyUtils.ParseCurrency(currencyToParse);
    }
}
