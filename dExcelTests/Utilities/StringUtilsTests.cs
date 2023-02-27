namespace dExcelTests.Utilities;

using NUnit.Framework;
using dExcel.Utilities;

[TestFixture]
public class StringUtilsTests
{
    [Test]
    public void RegexMatchTest()
    {
        const string expected = "C:\\Test\\Folder\\";
        string actual = StringUtils.RegexMatch("C:\\Test\\Folder\\File.txt", ".+\\\\");
        Assert.AreEqual(expected, actual);
    }
}
