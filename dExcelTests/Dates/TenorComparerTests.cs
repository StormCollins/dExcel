namespace dExcelTests.Dates;

using dExcel.Dates;
using NUnit.Framework;
using Omicron;

[TestFixture]
public class TenorComparerTests
{
    public static IEnumerable<TestCaseData> ComparerTestCaseData()
    {
        yield return new TestCaseData(null, null).Returns(0);
        yield return new TestCaseData(null, new Tenor(3, TenorUnit.Month)).Returns(-1);
        yield return new TestCaseData(new Tenor(3, TenorUnit.Month), null).Returns(1);
        yield return new TestCaseData(new Tenor(3, TenorUnit.Month), new Tenor(3, TenorUnit.Month)).Returns(0);
        yield return new TestCaseData(new Tenor(3, TenorUnit.Month), new Tenor(4, TenorUnit.Month)).Returns(-1);
        yield return new TestCaseData(new Tenor(4, TenorUnit.Month), new Tenor(3, TenorUnit.Month)).Returns(1);
        yield return new TestCaseData(new Tenor(1, TenorUnit.Day), new Tenor(1, TenorUnit.Week)).Returns(-1);
        yield return new TestCaseData(new Tenor(1, TenorUnit.Day), new Tenor(1, TenorUnit.Month)).Returns(-1);
        yield return new TestCaseData(new Tenor(1, TenorUnit.Day), new Tenor(1, TenorUnit.Year)).Returns(-1);
        yield return new TestCaseData(new Tenor(1, TenorUnit.Week), new Tenor(1, TenorUnit.Day)).Returns(1);
        yield return new TestCaseData(new Tenor(1, TenorUnit.Month), new Tenor(1, TenorUnit.Day)).Returns(1);
        yield return new TestCaseData(new Tenor(1, TenorUnit.Year), new Tenor(1, TenorUnit.Day)).Returns(1);
    }
    
    [Test]
    [TestCaseSource(nameof(ComparerTestCaseData))]
    public int CompareTest(Tenor? t1, Tenor? t2)
    {
        TenorComparer comparer = new();
        return comparer.Compare(t1, t2);
    }
}
