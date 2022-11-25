namespace dExcel.Mathematics;

using System;
using System.Linq;
using System.Numerics;
using ExcelDna.Integration;
using MathNet.Numerics;
using QLNet;
using mni = MathNet.Numerics.Interpolation;

/// <summary>
/// A collection of mathematical utility functions.
/// </summary>
public static class MathUtils
{
    [ExcelFunction(
        Name = "d.Math_Interpolate2D",
        Description = "Performs linear, exponential, or flat interpolation on a two-dimensional surface for two given points.",
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
            return $"{CommonUtils.DExcelErrorPrefix} Extrapolation not supported.";
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

    [ExcelFunction(
        Name = "d.Math_Interpolate",
        Description = "Performs linear, exponential, or flat interpolation on a range for a single given point.\n" +
                      "Deprecates AQS function: 'DT_Interp' and 'DT_Interp1'",
        Category = "∂Excel: Mathematics")]
    public static object Interpolate(
        [ExcelArgument(Name = "X-values", Description = "Independent variable.")]
        object[,] xValues,
        [ExcelArgument(Name = "Y-values", Description = "Dependent variable.")]
        object[,] yValues,
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
        // If xValues or yValues is a row, then transpose so that a column is supplied:
        object[,] Transpose(object[,] matrix)
        {
            int matrixRows = matrix.GetLength(0);
            int matrixCols = matrix.GetLength(1);

            object[,] matrixTransposed = new object[matrixCols, matrixRows];

            for (int i = 0; i < matrixRows; i++)
            {
                for (int j = 0; j < matrixCols; j++)
                {
                    matrixTransposed[j, i] = matrix[i, j];
                }
            }

            return matrixTransposed;
        }

        if (xValues.GetLength(1) > xValues.GetLength(0))
            xValues = Transpose(xValues);

        if (yValues.GetLength(1) > yValues.GetLength(0))
            yValues = Transpose(yValues);

        if ((xValues.GetLength(0) == yValues.GetLength(0)) && (xValues.GetLength(1)==1) && (yValues.GetLength(1) == 1))
        {
            int rowCount = Math.Max(xValues.GetLength(0), yValues.GetLength(0));
            double[] x = new double[rowCount];
            double[] y = new double[rowCount];
            for (int i = 0; i < rowCount; i++)
            {
                x[i] = (double)xValues[i, 0];
                y[i] = (double)yValues[i, 0];
            }
            mni.IInterpolation interpolator = null;

            switch (method.ToUpper())
            {
                case "LINEAR":
                    interpolator = mni.LinearSpline.Interpolate(x, y);
                    return interpolator.Interpolate(xi);
                case "EXPONENTIAL":
                    return ExponentialInterpolation();
                case "FLAT":
                    interpolator = mni.StepInterpolation.Interpolate(x, y);
                    return interpolator.Interpolate(xi);
                default:
                    return CommonUtils.DExcelErrorMessage("Invalid method of interpolation specified.");
            }

            // Log-linear interpolation fails for negative y-values therefore we move to the complex plane here then 
            // back to real numbers.
            double ExponentialInterpolation()
            {

                int index0 = Array.IndexOf(x, x.Where(element => element <= xi).Max());
                int index1 = Array.IndexOf(x, x.Where(element => element >= xi).Min());

                if (index0 == index1)
                {
                    return y[index0];
                }
                else
                {
                    Complex xiComplex = (Complex)xi;
                    Complex x0Complex = (Complex)x[index0];
                    Complex x1Complex = (Complex)x[index1];
                    Complex y0Complex = (Complex)y[index0];
                    Complex y1Complex = (Complex)y[index1];
                    Complex yi = (Complex.Log(y1Complex) - Complex.Log(y0Complex)) / (x1Complex - x0Complex) * (xiComplex - x0Complex) + Complex.Log(y0Complex);
                    Complex outputY = Complex.Exp(yi);
                    return (double)outputY.Real;
                }
            }
        }

        return CommonUtils.DExcelErrorMessage("Row dimensions do not match or there is more than one column in x or y.");
    }
}
