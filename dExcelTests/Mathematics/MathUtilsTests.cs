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
        object[,] xValues = { { 1.0 }, { 2.0 }, { 3.0 }, { 4.0 } };
        object[,] yValues = { { 2.0 }, { 4.0 }, { 6.0 }, { 8.0 } };

        double actual = (double)MathUtils.Interpolate(xValues, yValues, 1.5, "linear");
        double expected = 3;

        Assert.AreEqual(expected, actual);
    }
}
