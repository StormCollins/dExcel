using QLNet;

namespace dExcelTests;

using dExcel;
using NUnit.Framework;

[TestFixture]
public class CommonUtilsTests
{
    [Test]
    [TestCase("FOL", BusinessDayConvention.Following)]
    [TestCase("FOLLOWING", BusinessDayConvention.Following)]
    [TestCase("MODFOL", BusinessDayConvention.ModifiedFollowing)]
    [TestCase("MODIFIEDFOLLOWING", BusinessDayConvention.ModifiedFollowing)]
    [TestCase("MODIFIEDPRECEDING", BusinessDayConvention.ModifiedPreceding)]
    [TestCase("MODIFIEDPRECEDING", BusinessDayConvention.ModifiedPreceding)]
    [TestCase("PRECEDING", BusinessDayConvention.Preceding)]
    public void TestParseBusinessDayConvention(string x, BusinessDayConvention businessDayConvention)
    {
        Assert.AreEqual(CommonUtils.ParseBusinessDayConvention(x), businessDayConvention);
    }
} 