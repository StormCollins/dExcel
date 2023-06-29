using dExcel.Utilities;
using NUnit.Framework;

namespace dExcelTests.Utilities;

[TestFixture]
public class ExtensionTests
{
    [Test]
    public void SplitCamelCaseTest()
    {
        Assert.AreEqual("splitThisText".SplitCamelCase(), "Split This Text");
        string s = "something";
    }
}
