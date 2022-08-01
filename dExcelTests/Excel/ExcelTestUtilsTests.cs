namespace dExcelTests.Excel;

using NUnit.Framework;
using dExcel.ExcelUtils;
using ExcelDna.Integration;

[TestFixture]
public class ExcelTestUtilsTests
{
    [Test]
    [TestCase(1.5, 1.5, ExpectedResult = "OK")]
    [TestCase(1.5, 2.5, ExpectedResult = "ERROR")]
    [TestCase("TEST", "TEST", ExpectedResult = "OK")]
    [TestCase("APPLES", "ORANGES", ExpectedResult = "ERROR")]
    public string TestEqual(object a, object b)
    {
        return ExcelTestUtils.Equal(a, b);
    }

    [Test]
    [TestCase(true, ExpectedResult = "OK")]
    [TestCase(false, ExpectedResult = "ERROR")]
    public object TestIsTrue(object x)
    {
        return ExcelTestUtils.IsTrue(x);
    }

    [Test]
    [TestCase(false, ExpectedResult = "OK")]
    [TestCase(true, ExpectedResult = "ERROR")]
    public object TestIsFalse(object x)
    {
        return ExcelTestUtils.IsFalse(x);
    }

    [Test]
    [TestCase(1.5, 0.5, ExpectedResult = "OK")]
    [TestCase(1.5, 1.5, ExpectedResult = "ERROR")]
    [TestCase(1.5, 2.5, ExpectedResult = "ERROR")]
    public string TestGreaterThan(double a, double b)
    {
        return ExcelTestUtils.GreaterThan(a, b);
    }

    [Test]
    [TestCase(1.5, 0.5, ExpectedResult = "ERROR")]
    [TestCase(1.5, 1.5, ExpectedResult = "ERROR")]
    [TestCase(1.5, 2.5, ExpectedResult = "OK")]
    public string TestLessThan(double a, double b)
    { return ExcelTestUtils.LessThan(a, b); }

    public static IEnumerable<TestCaseData> AndTestCaseData()
    {
        yield return new TestCaseData("OK").Returns("OK");
        yield return new TestCaseData((object)new object[] { "OK", "OK" }).Returns("OK");
        yield return new TestCaseData(new object[,] { { "OK" }, { "OK" } }).Returns("OK");
        yield return new TestCaseData(new object[,] { { "OK", "OK" } }).Returns("OK");
        yield return new TestCaseData((object)new object[] { "OK" , "ERROR" }).Returns("ERROR");
        yield return new TestCaseData(new object[,] { { "OK" }, { "ERROR" } }).Returns("ERROR");
        yield return new TestCaseData(new object[,] { { "OK", "ERROR" } }).Returns("ERROR");
        yield return new TestCaseData((object)new object[] { "OK" , "WARNING" }).Returns("WARNING");
        yield return new TestCaseData(new object[,] { { "OK" }, { "WARNING" } }).Returns("WARNING");
        yield return new TestCaseData(new object[,] { { "OK", "WARNING" } }).Returns("WARNING");
        yield return new TestCaseData((object)new object[] { "WARNING" , "WARNING" }).Returns("WARNING");
        yield return new TestCaseData(new object[,] { { "WARNING" }, { "WARNING" } }).Returns("WARNING");
        yield return new TestCaseData(new object[,] { { "WARNING", "WARNING" } }).Returns("WARNING");
        yield return new TestCaseData((object)new object[] { "ERROR" , "WARNING" }).Returns("ERROR");
        yield return new TestCaseData(new object[,] { { "ERROR" }, { "WARNING" } }).Returns("ERROR");
        yield return new TestCaseData(new object[,] { { "ERROR", "WARNING" } }).Returns("ERROR");
        yield return new TestCaseData((object)new object[] { "ERROR" , "ERROR" }).Returns("ERROR");
        yield return new TestCaseData(new object[,] { { "ERROR" }, { "ERROR" } }).Returns("ERROR");
        yield return new TestCaseData(new object[,] { { "ERROR", "ERROR" } }).Returns("ERROR");
        yield return new TestCaseData(new object[,] { { "OK", ExcelError.ExcelErrorNA.ToString() } }).Returns("ERROR");
        yield return new TestCaseData(new object[,] { { "OK" }, { ExcelError.ExcelErrorNA.ToString() } }).Returns("ERROR");
        yield return new TestCaseData((object)new object[]{ "OK", ExcelError.ExcelErrorNA.ToString() }).Returns("ERROR");
    }

    [Test]
    [TestCaseSource(nameof(AndTestCaseData))]
    public string TestAnd(object range)
    {
        return ExcelTestUtils.And(range);
    }
    
    public static IEnumerable<TestCaseData> OrTestCaseData()
    {
        yield return new TestCaseData("OK").Returns("OK");
        yield return new TestCaseData((object)new object[] { "OK", "OK" }).Returns("OK");
        yield return new TestCaseData(new object[,] { { "OK" }, { "OK" } }).Returns("OK");
        yield return new TestCaseData(new object[,] { { "OK", "OK" } }).Returns("OK");
        yield return new TestCaseData((object)new object[] { "OK" , "ERROR" }).Returns("OK");
        yield return new TestCaseData(new object[,] { { "OK" }, { "ERROR" } }).Returns("OK");
        yield return new TestCaseData(new object[,] { { "OK", "ERROR" } }).Returns("OK");
        yield return new TestCaseData((object)new object[] { "OK" , "WARNING" }).Returns("OK");
        yield return new TestCaseData(new object[,] { { "OK" }, { "WARNING" } }).Returns("OK");
        yield return new TestCaseData(new object[,] { { "OK", "WARNING" } }).Returns("OK");
        yield return new TestCaseData((object)new object[] { "WARNING" , "WARNING" }).Returns("WARNING");
        yield return new TestCaseData(new object[,] { { "WARNING" }, { "WARNING" } }).Returns("WARNING");
        yield return new TestCaseData(new object[,] { { "WARNING", "WARNING" } }).Returns("WARNING");
        yield return new TestCaseData((object)new object[] { "ERROR" , "WARNING" }).Returns("WARNING");
        yield return new TestCaseData(new object[,] { { "ERROR" }, { "WARNING" } }).Returns("WARNING");
        yield return new TestCaseData(new object[,] { { "ERROR", "WARNING" } }).Returns("WARNING");
        yield return new TestCaseData((object)new object[] { "ERROR" , "ERROR" }).Returns("ERROR");
        yield return new TestCaseData(new object[,] { { "ERROR" }, { "ERROR" } }).Returns("ERROR");
        yield return new TestCaseData(new object[,] { { "ERROR", "ERROR" } }).Returns("ERROR");
        yield return new TestCaseData(new object[,] { { "OK", ExcelError.ExcelErrorNA.ToString() } }).Returns("ERROR");
        yield return new TestCaseData(new object[,] { { "OK" }, { ExcelError.ExcelErrorNA.ToString() } }).Returns("ERROR");
        yield return new TestCaseData((object)new object[]{ "OK", ExcelError.ExcelErrorNA.ToString() }).Returns("ERROR");
    }
    
    [Test]
    [TestCaseSource(nameof(OrTestCaseData))]
    public string TestOr(object range)
    {
        return ExcelTestUtils.Or(range);
    }
    
    public static IEnumerable<TestCaseData> NotTestCaseData()
    {
        yield return new TestCaseData(new object[,] { { "OK" }, { "ERROR" }, { "WARNING" } }).Returns(new object[,] { { "ERROR" }, { "OK" }, { "WARNING" } });
        yield return new TestCaseData(new object[,] { { "OK" , "ERROR", "WARNING" } }).Returns(new object[,] { { "ERROR", "OK", "WARNING" } } );
        yield return new TestCaseData(new object[,] { { "OK" , "ERROR", ExcelError.ExcelErrorNA.ToString() } }).Returns(new object[,] { { "ERROR", "ERROR", "ERROR" } });
    }

    [Test]
    [TestCaseSource(nameof(NotTestCaseData))]
    public object TestNot(object[,] range)
    {
        return ExcelTestUtils.Not(range);
    }
}
