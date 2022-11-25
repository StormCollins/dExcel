namespace dExcelTests.Mathematics;

using NUnit.Framework;
using dExcel.Mathematics;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

[TestFixture]
public class MathUtilsTests
{
    [Test]
    public void TestLinearInterpolation()
    {
        object[,] xValues = new object[,] { { 1 }, { 2 }, { 3 }, { 4 } };
        object[,] yValues = new object[,] { { 2 }, { 4 }, { 6 }, { 8 } };

        double actual = (double)MathUtils.InterpolateTwoColumns(xValues, yValues, 1.5, "l");
        double expected = 3;

        Assert.AreEqual(expected, actual);
    }
}
