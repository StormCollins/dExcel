using NUnit.Framework;
using QL = QuantLib;

namespace dExcelTests.Currencies;

using dExcel.FX;

[TestFixture]
public sealed class CurrencyUtilsTests
{
    public static IEnumerable<TestCaseData> CurrencyTestData()
    {
        yield return new TestCaseData("EUR").Returns(new QL.EURCurrency().code());
        yield return new TestCaseData("GBP").Returns(new QL.GBPCurrency().code());
        yield return new TestCaseData("USD").Returns(new QL.USDCurrency().code());
        yield return new TestCaseData("ZAR").Returns(new QL.ZARCurrency().code());
        yield return new TestCaseData("Invalid").Returns(null);
    }
    
    [Test]
    [TestCaseSource(nameof(CurrencyTestData))]
    public string? TestParseCurrency(string currencyToParse)
    {
        return CurrencyUtils.ParseCurrency(currencyToParse)?.code();
    }
}
