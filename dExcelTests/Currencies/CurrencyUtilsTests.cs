using dExcel.Currencies;
using NUnit.Framework;
using QL = QuantLib;

namespace dExcelTests.Currencies;

[TestFixture]
public sealed class CurrencyUtilsTests
{
    public static IEnumerable<TestCaseData> CurrencyTestData()
    {
        yield return new TestCaseData("EUR").Returns(new QL.EURCurrency());
        yield return new TestCaseData("GBP").Returns(new QL.GBPCurrency());
        yield return new TestCaseData("USD").Returns(new QL.USDCurrency());
        yield return new TestCaseData("ZAR").Returns(new QL.ZARCurrency());
        yield return new TestCaseData("Invalid").Returns(null);
    }
    
    [Test]
    [TestCaseSource(nameof(CurrencyTestData))]
    public QL.Currency? TestParseCurrency(string currencyToParse)
    {
        return CurrencyUtils.ParseCurrency(currencyToParse);
    }
}
