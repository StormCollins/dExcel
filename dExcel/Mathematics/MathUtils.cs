namespace dExcel.Mathematics;

using System;
using System.Linq;
using System.Numerics;
using ExcelDna.Integration;
using ExcelUtils;
using Utilities;
using mni = MathNet.Numerics.Interpolation;

/// <summary>
/// A collection of mathematical utility functions.
/// </summary>
public static class MathUtils
{
    /// <summary>
    /// Performs linear, exponential, or flat interpolation on a 2D surface for a given (x,y) coordinate.
    /// </summary>
    /// <param name="xy">Matrix from which to interpolate, where X is the horizontal dimension and Y the vertical
    /// dimension. XY must include the numeric row and column headings.</param>
    /// <param name="x">X-value (along horizontal-axis) for which to interpolate.</param>
    /// <param name="y">Y-value (along vertical-axis) for which to interpolate.</param>
    /// <param name="method">Method of interpolation: 'linear', 'exponential', 'flat'</param>
    /// <returns></returns>
    [ExcelFunction(
        Name = "d.Math_Interpolate2D",
        Description = "Performs linear, exponential, or flat interpolation on a 2D surface for a given (x,y) coordinate.",
        Category = "∂Excel: Mathematics")]
    public static object Interpolate2D(
        [ExcelArgument(Name = "XY", Description = "Matrix from which to interpolate, where X is the horizontal dimension and Y the vertical dimension. XY must include the numeric row and column headings.")]
        object[,] xy,
        [ExcelArgument(Name = "X", Description = "X-value (along horizontal-axis) for which to interpolate.")]
        double x,
        [ExcelArgument(Name = "Y", Description = "Y-value (along vertical-axis) for which to interpolate.")]
        double y,
        [ExcelArgument(
            Name = "Method",
            Description = "Method of interpolation: 'linear', 'exponential', 'flat'")]
        string method)
    {
#if DEBUG
        CommonUtils.InFunctionWizard();
#endif
        // TODO: Add check that 'xy' is numeric data only.
        int xCount = xy.GetLength(1) - 1;
        int yCount = xy.GetLength(0) - 1;
        double[] xValues = new double[xCount];
        object[,] xValuesObject = new object[1,xCount];
        double[] yValues = new double[yCount];
        object[,] yValuesObject = new object[yCount, 1];

        for (int i = 0; i < xCount; i++)
        {
            xValues[i] = (double)xy[0, i + 1];
            xValuesObject[0,i] = xy[0, i + 1];
        }

        for (int i = 0; i < yCount; i++)
        {
            yValues[i] = (double)xy[i + 1, 0];
            yValuesObject[i,0] = xy[i + 1, 0];
        }

        if (x < xValues.Min() || x > xValues.Max() || y < yValues.Min() || y > yValues.Max())
        {
            return CommonUtils.DExcelErrorMessage("Extrapolation not supported.");
        }

        double yLeft = yValues.Where(element => element <= y).Max();
        int yIndexLeft = Array.IndexOf(yValues, yLeft);

        double yRight = yValues.Where(element => element >= y).Min();
        int yIndexRight = Array.IndexOf(yValues, yRight);

        object[,] zLeftRange = new object[1, xCount];
        for (int i = 0; i < xCount; i++)
        {
            zLeftRange[0, i] = xy[yIndexLeft + 1, i + 1];
        }

        object[,] zRightRange = new object[1, xCount];
        for (int i = 0; i < xCount; i++)
        {
            zRightRange[0, i] = xy[yIndexRight + 1, i + 1];
        }

        object zLeft = Interpolate(xValuesObject, zLeftRange, x, method);
        object zRight = Interpolate(xValuesObject, zRightRange, x, method);

        object z;

        if (zLeft.Equals(zRight))
        {
            z = zLeft;
        }
        else
        {
            object[,] zLeftAndRight = { { zLeft }, { zRight } };
            object[,] yLeftAndRight = { { yLeft }, { yRight } };

            z = Interpolate(yLeftAndRight, zLeftAndRight, y, method);
        }
        
        return z;
    }

    /// <summary>
    /// Performs linear, exponential, or flat interpolation on a range for a single given point.
    /// </summary>
    /// <param name="xRange">Independent variable.</param>
    /// <param name="yRange">Dependent variable.</param>
    /// <param name="xi">Value for which to interpolate.</param>
    /// <param name="method">Method of interpolation: 'linear', 'exponential', 'flat'</param>
    /// <returns>Interpolated y-value.</returns>
    [ExcelFunction(
        Name = "d.Math_Interpolate",
        Description = "Performs linear, exponential, or flat interpolation on a range for a single given point.\n" +
                      "Deprecates AQS functions: 'DT_Interp' and 'DT_Interp1'",
        Category = "∂Excel: Mathematics")]
    public static object Interpolate(
        [ExcelArgument(Name = "X-values", Description = "Independent variable.")]
        object[,] xRange,
        [ExcelArgument(Name = "Y-values", Description = "Dependent variable.")]
        object[,] yRange,
        [ExcelArgument(Name = "Xi", Description = "Value for which to interpolate.")]
        double xi,
        [ExcelArgument(
            Name = "Method",
            Description = "Method of interpolation: 'linear', 'exponential', 'flat'")]
        string method)
    {
#if DEBUG
        CommonUtils.InFunctionWizard();
#endif
        if (xRange.GetLength(0) > 1 && xRange.GetLength(1) > 1)
        {
            return CommonUtils.DExcelErrorMessage("x-value range has too many dimensions.");
        }
        
        if (yRange.GetLength(0) > 1 && yRange.GetLength(1) > 1)
        {
            return CommonUtils.DExcelErrorMessage("y-value range has too many dimensions.");
        }
        
        List<double> xValues = ExcelArrayUtils.ConvertExcelRangeToList<double>(xRange);
        List<double> yValues = ExcelArrayUtils.ConvertExcelRangeToList<double>(yRange);

        if (xValues.Distinct().Count() != xValues.Count)
        {
            return CommonUtils.DExcelErrorMessage("Duplicate values in x-values range.");
        }
        
        if (xValues.Count != yValues.Count)
        {
            return CommonUtils.DExcelErrorMessage("Dimensions of x and y ranges don't match.");
        }
       
        mni.IInterpolation? interpolator = null;

        switch (method.ToUpper())
        {
            case "LINEAR":
                interpolator = mni.LinearSpline.Interpolate(xValues, yValues);
                return interpolator.Interpolate(xi);
            case "EXPONENTIAL":
                return ExponentialInterpolation();
            case "FLAT":
                interpolator = mni.StepInterpolation.Interpolate(xValues, yValues);
                return interpolator.Interpolate(xi);
            default:
                return CommonUtils.DExcelErrorMessage($"Unsupported interpolation: '{method}'");
        }

        // Log-linear interpolation fails for negative y-values therefore we move to the complex plane here then 
        // back to real numbers.
        double ExponentialInterpolation()
        {
            int lowerXIndex = xValues.IndexOf(xValues.Where(x => x <= xi).Max());
            int upperXIndex = xValues.IndexOf(xValues.Where(x => x >= xi).Min());

            if (lowerXIndex == upperXIndex)
            {
                return yValues[lowerXIndex];
            }

            Complex xiComplex = xi;
            Complex x0Complex = xValues[lowerXIndex];
            Complex x1Complex = xValues[upperXIndex];
            Complex y0Complex = yValues[lowerXIndex];
            Complex y1Complex = yValues[upperXIndex];
            Complex yi = (Complex.Log(y1Complex) - Complex.Log(y0Complex)) / (x1Complex - x0Complex) * (xiComplex - x0Complex) + Complex.Log(y0Complex);
            Complex outputY = Complex.Exp(yi);
            return outputY.Real;
        }   
    }
}
