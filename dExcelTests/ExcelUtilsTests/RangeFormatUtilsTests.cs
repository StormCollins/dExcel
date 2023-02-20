namespace dExcelTests.ExcelUtilsTests;

using dExcel.ExcelUtils;
using NUnit.Framework;

[TestFixture]
public class RangeFormatUtilsTests
{
    [Test]
    [TestCase(1, "A")]
    [TestCase(27, "AA")]
    [TestCase(703, "AAA")]
    public void GetColumnLetterTest(int columnNumber, string columnLetter)
    {
        Assert.AreEqual(RangeFormatUtils.GetColumnLetter(columnNumber), columnLetter);
    }
}
