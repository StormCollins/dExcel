namespace dExcelTests.Mathematics;

using NUnit.Framework;
using dExcel.Mathematics;
using dExcel.Utilities;

[TestFixture]
public class MathUtilsTests
{
    [Test]
    public void LinearInterpolationTest()
    {
        object[,] xValues = { { 1.0 }, { 2.0 }, { 3.0 }, { 4.0 } };
        object[,] yValues = { { 2.0 }, { 4.0 }, { 6.0 }, { 8.0 } };
        double actual = (double)MathUtils.Interpolate(xValues, yValues, 1.5, "linear");
        const double expected = 3;
        Assert.AreEqual(expected, actual);
    }
    
    [Test]
    public void LinearInterpolationOnRowsTest()
    {
        object[,] xValues = { { 1.0, 2.0, 3.0, 4.0 } };
        object[,] yValues = { { 2.0, 4.0, 6.0, 8.0 } };
        double actual = (double)MathUtils.Interpolate(xValues, yValues, 1.5, "linear");
        const double expected = 3;
        Assert.AreEqual(expected, actual);
    }

    [Test]
    public void ExponentialInterpolationTest()
    {
        object[,] xValues = { { 1.0, 2.0, 3.0, 4.0 } };
        object[,] yValues = { { 2.0, 4.0, 6.0, 8.0 } };
        double actual = (double)MathUtils.Interpolate(xValues, yValues, 1.5, "exponential");
        double expected = Math.Exp((Math.Log(4.0) - Math.Log(2.0)) / (2.0 - 1.0) * (1.5 - 1.0) + Math.Log(2.0));
        Assert.AreEqual(expected, actual);
    }
    
    [Test]
    public void ExponentialInterpolationOnNodeTest()
    {
        object[,] xValues = { { 1.0, 2.0, 3.0, 4.0 } };
        object[,] yValues = { { 2.0, 4.0, 6.0, 8.0 } };
        double actual = (double)MathUtils.Interpolate(xValues, yValues, 2, "exponential");
        const double expected = 4.0;
        Assert.AreEqual(expected, actual);
    }

    [Test]
    public void FlatInterpolationTest()
    {
        object[,] xValues = { { 1.0, 2.0, 3.0, 4.0 } };
        object[,] yValues = { { 2.0, 4.0, 6.0, 8.0 } };
        double actual = (double)MathUtils.Interpolate(xValues, yValues, 1.5, "flat");
        const double expected = 2.0;
        Assert.AreEqual(expected, actual);
    }

    [Test]
    public void DuplicateXValuesTest()
    {
        object[,] xValues = { { 1.0, 1.0, 3.0, 4.0 } };
        object[,] yValues = { { 2.0, 4.0, 6.0, 8.0 } };
        string actual = MathUtils.Interpolate(xValues, yValues, 1.5, "flat").ToString();
        string expected = CommonUtils.DExcelErrorMessage("Duplicate values in x-values range.");
        Assert.AreEqual(expected, actual);
    }

    [Test]
    public void UnsupportedInterpolationTest()
    {
        object[,] xValues = { { 1.0, 2.0, 3.0, 4.0 } };
        object[,] yValues = { { 2.0, 4.0, 6.0, 8.0 } };
        string? actual = MathUtils.Interpolate(xValues, yValues, 1.5, "INVALID").ToString();
        string expected = CommonUtils.DExcelErrorMessage("Unsupported interpolation: 'INVALID'");
        Assert.AreEqual(expected, actual);
    }
    
    [Test]
    public void MismatchingDimensionsTest()
    {
        object[,] xValues = { { 1.0 }, { 2.0 }, { 3.0 }, { 4.0 }, { 5.0 }};
        object[,] yValues = { { 2.0 }, { 4.0 }, { 6.0 }, { 8.0 } };
        string? actual = MathUtils.Interpolate(xValues, yValues, 1.5, "linear").ToString();
        string expected = CommonUtils.DExcelErrorMessage("Dimensions of x and y ranges don't match.");
        Assert.AreEqual(expected, actual);
    }

    [Test]
    public void TooManyXValueRangeDimensionsTest()
    {
        object[,] xValues = { { 1.0, 5.0 }, { 2.0, 6.0 }, { 3.0, 7.0 }, { 4.0, 8.0 }};
        object[,] yValues = { { 2.0 }, { 4.0 }, { 6.0 }, { 8.0 } };
        string? actual = MathUtils.Interpolate(xValues, yValues, 1.5, "linear").ToString();
        string expected = CommonUtils.DExcelErrorMessage("x-value range has too many dimensions.");
        Assert.AreEqual(expected, actual);
    }
    
    [Test]
    public void TooManyYValueRangeDimensionsTest()
    {
        object[,] xValues = { { 1.0 }, { 2.0 }, { 3.0 }, { 4.0 }};
        object[,] yValues = { { 2.0, 9.0 }, { 4.0, 10.0 }, { 6.0, 11.0 }, { 8.0, 12.0 } };
        string? actual = MathUtils.Interpolate(xValues, yValues, 1.5, "linear").ToString();
        string expected = CommonUtils.DExcelErrorMessage("y-value range has too many dimensions.");
        Assert.AreEqual(expected, actual);
    }
    
    [Test]
    public void BilinearInterpolateTest()
    {
        object[,] xy =
        {
            {" ",  0.0,  1.0,  2.0,  3.0},
            {0.0,  1.0,  2.0,  3.0,  4.0},
            {1.0,  5.0,  6.0,  7.0,  8.0},
            {2.0,  9.0, 10.0, 11.0, 12.0},
            {3.0, 13.0, 14.0, 15.0, 16.0},
            {4.0, 17.0, 18.0, 19.0, 20.0}
        };
        
        Assert.AreEqual(9.5, MathUtils.Interpolate2D(xy, 2.5, 1.5, "linear"));
    }

    [Test]
    public void BilinearInterpolateOnEdgeTest()
    {
        object[,] xy =
        {
            {" ",  0.0,  1.0,  2.0,  3.0},
            {0.0,  1.0,  2.0,  3.0,  4.0},
            {1.0,  5.0,  6.0,  7.0,  8.0},
            {2.0,  9.0, 10.0, 11.0, 12.0},
            {3.0, 13.0, 14.0, 15.0, 16.0},
            {4.0, 17.0, 18.0, 19.0, 20.0}
        };
        
        Assert.AreEqual(7.0, MathUtils.Interpolate2D(xy, 2, 1, "linear"));
    }
    
    [Test]
    public void BilinearExtrapolationNotSupportedTest()
    {
        object[,] xy =
        {
            {" ",  0.0,  1.0,  2.0,  3.0},
            {0.0,  1.0,  2.0,  3.0,  4.0},
            {1.0,  5.0,  6.0,  7.0,  8.0},
            {2.0,  9.0, 10.0, 11.0, 12.0},
            {3.0, 13.0, 14.0, 15.0, 16.0},
            {4.0, 17.0, 18.0, 19.0, 20.0}
        };

        string expected = CommonUtils.DExcelErrorMessage("Extrapolation not supported.");
        Assert.AreEqual(expected, MathUtils.Interpolate2D(xy, 9.5, 1.5, "linear"));
    }
}
