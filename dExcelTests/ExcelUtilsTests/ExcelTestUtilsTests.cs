namespace dExcelTests.ExcelUtilsTests;

using dExcel.ExcelUtils;
using ExcelDna.Integration;
using NUnit.Framework;

[TestFixture]
public class ExcelTestUtilsTests
{
    [Test]
    [TestCase(1.5, 1.5, ExpectedResult = "Ok")]
    [TestCase(1.5, 2.5, ExpectedResult = "Error")]
    [TestCase("TEST", "TEST", ExpectedResult = "Ok")]
    [TestCase("APPLES", "ORANGES", ExpectedResult = "Error")]
    [TestCase("ExcelErrorNA", "ORANGES", ExpectedResult = "Error")]
    public string TestEqual(object a, object b)
    {
        return ExcelTestUtils.Equal(a, b);
    }

    [Test]
    [TestCase(true, ExpectedResult = "Ok")]
    [TestCase(false, ExpectedResult = "Error")]
    [TestCase("ExcelErrorNA", ExpectedResult = "Error")]
    public object TestIsTrue(object x)
    {
        return ExcelTestUtils.IsTrue(x);
    }

    [Test]
    [TestCase(false, ExpectedResult = "Ok")]
    [TestCase(true, ExpectedResult = "Error")]
    [TestCase("ExcelErrorNA", ExpectedResult = "Error")]
    public object TestIsFalse(object x)
    {
        return ExcelTestUtils.IsFalse(x);
    }

    [Test]
    [TestCase(1.5, 0.5, ExpectedResult = "Ok")]
    [TestCase(1.5, 1.5, ExpectedResult = "Error")]
    [TestCase(1.5, 2.5, ExpectedResult = "Error")]
    [TestCase("ExcelErrorNA", 2.5, ExpectedResult = "Error")]
    public string TestGreaterThan(object a, object b)
    {
        return ExcelTestUtils.GreaterThan(a, b);
    }

    [Test]
    [TestCase(1.5, 0.5, ExpectedResult = "Error")]
    [TestCase(1.5, 1.5, ExpectedResult = "Error")]
    [TestCase(1.5, 2.5, ExpectedResult = "Ok")]
    [TestCase("ExcelErrorNA", 2.5, ExpectedResult = "Error")]
    public string TestLessThan(object a, object b)
    {
        return ExcelTestUtils.LessThan(a, b);
    }

    public static IEnumerable<TestCaseData> AndTestCaseData()
    {
        yield return new TestCaseData("Ok").Returns("Ok");
        yield return new TestCaseData(ExcelError.ExcelErrorNA.ToString()).Returns("Error");
        yield return new TestCaseData((object)new object[] { "Ok", "Ok" }).Returns("Ok");
        yield return new TestCaseData(new object[,] { { "Ok" }, { "Ok" } }).Returns("Ok");
        yield return new TestCaseData(new object[,] { { "Ok", "Ok" } }).Returns("Ok");
        yield return new TestCaseData((object)new object[] { "Ok" , "Error" }).Returns("Error");
        yield return new TestCaseData(new object[,] { { "Ok" }, { "Error" } }).Returns("Error");
        yield return new TestCaseData(new object[,] { { "Ok", "Error" } }).Returns("Error");
        yield return new TestCaseData((object)new object[] { "Ok" , "Warning" }).Returns("Warning");
        yield return new TestCaseData(new object[,] { { "Ok" }, { "Warning" } }).Returns("Warning");
        yield return new TestCaseData(new object[,] { { "Ok", "Warning" } }).Returns("Warning");
        yield return new TestCaseData((object)new object[] { "Warning" , "Warning" }).Returns("Warning");
        yield return new TestCaseData(new object[,] { { "Warning" }, { "Warning" } }).Returns("Warning");
        yield return new TestCaseData(new object[,] { { "Warning", "Warning" } }).Returns("Warning");
        yield return new TestCaseData((object)new object[] { "Error" , "Warning" }).Returns("Error");
        yield return new TestCaseData(new object[,] { { "Error" }, { "Warning" } }).Returns("Error");
        yield return new TestCaseData(new object[,] { { "Error", "Warning" } }).Returns("Error");
        yield return new TestCaseData((object)new object[] { "Error" , "Error" }).Returns("Error");
        yield return new TestCaseData(new object[,] { { "Error" }, { "Error" } }).Returns("Error");
        yield return new TestCaseData(new object[,] { { "Error", "Error" } }).Returns("Error");
        yield return new TestCaseData((object)new object[]{ "Ok", ExcelError.ExcelErrorNA.ToString() }).Returns("Error");
        yield return new TestCaseData(new object[,] { { "Ok", ExcelError.ExcelErrorNA.ToString() } }).Returns("Error");
        yield return new TestCaseData(new object[,] { { "Ok" }, { ExcelError.ExcelErrorNA.ToString() } }).Returns("Error");
    }

    [Test]
    [TestCaseSource(nameof(AndTestCaseData))]
    public string TestAnd(object range)
    {
        return ExcelTestUtils.And(range);
    }
    
    public static IEnumerable<TestCaseData> OrTestCaseData()
    {
        yield return new TestCaseData("Ok").Returns("Ok");
        yield return new TestCaseData(ExcelError.ExcelErrorNA.ToString().ToUpper()).Returns("Error");
        yield return new TestCaseData((object)new object[] { "Ok", "Ok" }).Returns("Ok");
        yield return new TestCaseData(new object[,] { { "Ok" }, { "Ok" } }).Returns("Ok");
        yield return new TestCaseData(new object[,] { { "Ok", "Ok" } }).Returns("Ok");
        yield return new TestCaseData((object)new object[] { "Ok" , "Error" }).Returns("Ok");
        yield return new TestCaseData(new object[,] { { "Ok" }, { "Error" } }).Returns("Ok");
        yield return new TestCaseData(new object[,] { { "Ok", "Error" } }).Returns("Ok");
        yield return new TestCaseData((object)new object[] { "Ok" , "Warning" }).Returns("Ok");
        yield return new TestCaseData(new object[,] { { "Ok" }, { "Warning" } }).Returns("Ok");
        yield return new TestCaseData(new object[,] { { "Ok", "Warning" } }).Returns("Ok");
        yield return new TestCaseData((object)new object[] { "Warning" , "Warning" }).Returns("Warning");
        yield return new TestCaseData(new object[,] { { "Warning" }, { "Warning" } }).Returns("Warning");
        yield return new TestCaseData(new object[,] { { "Warning", "Warning" } }).Returns("Warning");
        yield return new TestCaseData((object)new object[] { "Error" , "Warning" }).Returns("Warning");
        yield return new TestCaseData(new object[,] { { "Error" }, { "Warning" } }).Returns("Warning");
        yield return new TestCaseData(new object[,] { { "Error", "Warning" } }).Returns("Warning");
        yield return new TestCaseData((object)new object[] { "Error" , "Error" }).Returns("Error");
        yield return new TestCaseData(new object[,] { { "Error" }, { "Error" } }).Returns("Error");
        yield return new TestCaseData(new object[,] { { "Error", "Error" } }).Returns("Error");
        yield return new TestCaseData((object)new object[]{ "Ok", ExcelError.ExcelErrorNA.ToString() }).Returns("Error");
        yield return new TestCaseData(new object[,] { { "Ok", ExcelError.ExcelErrorNA.ToString() } }).Returns("Error");
        yield return new TestCaseData(new object[,] { { "Ok" }, { ExcelError.ExcelErrorNA.ToString() } }).Returns("Error");
    }
    
    [Test]
    [TestCaseSource(nameof(OrTestCaseData))]
    public string TestOr(object range)
    {
        return ExcelTestUtils.Or(range);
    }
    
    public static IEnumerable<TestCaseData> NotTestCaseData()
    {
        yield return new TestCaseData(new object[,] { { "Ok" }, { "Error" }, { "Warning" } }).Returns(new object[,] { { "Error" }, { "Ok" }, { "Warning" } });
        yield return new TestCaseData(new object[,] { { "Ok" , "Error", "Warning" } }).Returns(new object[,] { { "Error", "Ok", "Warning" } } );
        yield return new TestCaseData(new object[,] { { "Ok" , "Error", ExcelError.ExcelErrorNA.ToString() } }).Returns(new object[,] { { "Error", "Error", "Error" } });
    }

    [Test]
    [TestCaseSource(nameof(NotTestCaseData))]
    public object TestNot(object[,] range)
    {
        return ExcelTestUtils.Not(range);
    }
}
