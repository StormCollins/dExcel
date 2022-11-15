namespace dExcel;

using System;
using System.Linq;
using ExcelDna.Integration;
using mni = MathNet.Numerics.Interpolation;
using MathNet.Numerics;

/// <summary>
/// A collection of mathematical utility functions.
/// </summary>
public static class MathUtils
{
    [ExcelFunction(
        Name = "d.Math_Bilinterp",
        Description = "Performs bi-linear interpolation on two variables.",
        Category = "∂Excel: Mathematics")]
    public static object Bilinterp(
        [ExcelArgument(Name = "XY", Description = "Matrix from which to interpolate.")]
        object[,] xy,
        [ExcelArgument(Name = "X", Description = "X-value (along column-axis) for which to interpolate.")]
        double x,
        [ExcelArgument(Name = "Y", Description = "Y-value (along row-axis) for which to interpolate.")]
        double y)
    {
#if DEBUG
        CommonUtils.InFunctionWizard();
#endif

        // TODO: Add check that 'xy' is numeric data only.
        var rowCount = xy.GetLength(0) - 1;
        var colCount = xy.GetLength(1) - 1;
        var rowValues = new double[rowCount];
        var colValues = new double[colCount];

        for (int i = 1; i <= rowCount; i++)
        {
            rowValues[i-1] = (double)xy[i, 0];
        }

        for (int i = 1; i <= colCount; i++)
        {
            colValues[i-1] = (double)xy[0, i];
        }

        if (y <= rowValues.Min() || y >= rowValues.Max() || x <= colValues.Min() || x >= colValues.Max())
        {
            return "ERROR: Extrapolation not supported.";
        }

        double z = 0.0;
        // TODO: Find the row values and column value satisfying this first.
        for (int i = 0; i < rowValues.Length - 1; i++)
        {
            for (int j = 0; j < colValues.Length - 1; j++)
            {
                if (y >= rowValues[i] && y < rowValues[i + 1] && x >= colValues[j] && x < colValues[j + 1])
                {
                    z =
                        (double)xy[i + 1, j + 1] * (rowValues[i + 1] - y) * (colValues[j + 1] - x) / (rowValues[i + 1] - rowValues[i]) / (colValues[j + 1] - colValues[j]) +
                        (double)xy[i + 2, j + 1] * (y - rowValues[i]) * (colValues[j + 1] - x) / (rowValues[i + 1] - rowValues[i]) / (colValues[j + 1] - colValues[j]) +
                        (double)xy[i + 1, j + 2] * (rowValues[i + 1] - y) * (x - colValues[j]) / (rowValues[i + 1] - rowValues[i]) / (colValues[j + 1] - colValues[j]) +
                        (double)xy[i + 2, j + 2] * (y - rowValues[i]) * (x - colValues[j]) / (rowValues[i + 1] - rowValues[i]) / (colValues[j + 1] - colValues[j]);
                }
            }
        }
        return z;
    }

    [ExcelFunction(
        Name = "d.Math_InterpolateContiguousArea",
        Description = "Performs linear or log-linear interpolation on a range for a single point.\n" +
                      "Deprecates AQS function: 'dt_interp1'",
        Category = "∂Excel: Mathematics")]
    public static object InterpolateContiguousArea(
        [ExcelArgument(
            Name = "XY",
            Description = "Contiguous region of data for interpolation. Assumes zeroth column are x values.")]
        object[,] xy,
        [ExcelArgument(Name = "Y-column", Description = "Integer specifying Y-column.")]
        int yColumn,
        [ExcelArgument(
            Name = "xi",
            Description = "Value for which to interpolate.")]
        double xi,
        [ExcelArgument(
            Name = "Method",
            Description = "Method of interpolation: 'l' = 'linear', 'e' = 'exponential/log linear' ")]
        string method)
    {
#if DEBUG
        CommonUtils.InFunctionWizard();
#endif
        var rowCount = Math.Max(xy.GetLength(0), xy.GetLength(1));
        var x = new double[rowCount];
        var y = new double[rowCount];
        for (int i = 0; i < rowCount; i++)
        {
            x[i] = (double)xy[i, 0];
            y[i] = (double)xy[i, yColumn - 1];
        }
        mni.IInterpolation interpolator = null;

        switch (method.ToUpper())
        {
            case "L":
                interpolator = mni.LinearSpline.Interpolate(x, y);
                break;
            case "E":
                interpolator = mni.LogLinear.Interpolate(x, y);
                break;
            default:
                break;
        }

        int index = 0;
        for (int i = 0; i < rowCount; i++)
        {
            if (x[i] <= xi && xi < x[i + 1])
            {
                index = i;
                break;
            }
        }

        Complex32 xiComplex = (Complex32)xi;
        Complex32 x0Complex = (Complex32)x[index];
        Complex32 x1Complex = (Complex32)x[index + 1];
        Complex32 y0Complex = (Complex32)y[index];
        Complex32 y1Complex = (Complex32)y[index + 1];

        Complex32 yi = (Complex32.Log(y1Complex) - Complex32.Log(y0Complex)) / (x1Complex - x0Complex) * (xiComplex - x0Complex) + Complex32.Log(y0Complex);
        Complex32 outputY = Complex32.Exp(yi);


        if (x.Min() <= xi && xi <= x.Max())
            return interpolator.Interpolate(xi);
        else
            return "Extrapolation not supported.";
    }

    [ExcelFunction(
        Name = "d.Math_InterpolateTwoColumns",
        Description = "Performs linear or log-linear interpolation on a range for a single point.\n" +
                      "Deprecates AQS function: 'dt_interp'",
        Category = "∂Excel: Mathematics")]
    public static object InterpolateTwoColumns(
        [ExcelArgument(Name = "X-column", Description = "Independent variable.")]
        object[,] xCol,
        [ExcelArgument(Name = "Y-column", Description = "Dependent Variable.")]
        object[,] yCol,
        [ExcelArgument(Name = "xi", Description = "Value for which to interpolate.")]
        double xi,
        [ExcelArgument(
            Name = "Method",
            Description = "Method of interpolation: 'l' = 'linear', 'e' = 'exponential/log linear' ")]
        string method)
    {
#if DEBUG
        CommonUtils.InFunctionWizard();
#endif

        if ((xCol.GetLength(0) == yCol.GetLength(0)) && (xCol.GetLength(1)==1) && (yCol.GetLength(1) == 1))
        {
            var rowCount = Math.Max(xCol.GetLength(0), yCol.GetLength(0));
            var x = new double[rowCount];
            var y = new double[rowCount];
            for (int i = 0; i < rowCount; i++)
            {
                x[i] = (double)xCol[i, 0];
                y[i] = (double)yCol[i, 0];
            }
            mni.IInterpolation interpolator = null;

            switch (method.ToUpper())
            {
                case "L":
                    interpolator = mni.LinearSpline.Interpolate(x, y);
                    break;
                case "E":
                    interpolator = mni.LogLinear.Interpolate(x, y);
                    break;
                default:
                    break;
            }

            int index = 0;
            for (int i = 0; i < rowCount; i++)
            {
                if (x[i] <= xi && xi < x[i + 1])
                {
                    index = i;
                    break;
                }
            }

            Complex32 xiComplex = (Complex32)xi;
            Complex32 x0Complex = (Complex32)x[index];
            Complex32 x1Complex = (Complex32)x[index + 1];
            Complex32 y0Complex = (Complex32)y[index];
            Complex32 y1Complex = (Complex32)y[index + 1];
            Complex32 yi = (Complex32.Log(y1Complex) - Complex32.Log(y0Complex)) / (x1Complex - x0Complex) * (xiComplex - x0Complex) + Complex32.Log(y0Complex);
            Complex32 outputY = Complex32.Exp(yi);
            return (double)outputY.Real;
        }
        else
        {
           return "ERROR: Row dimensions do not match or there is more than one column in x or y.";
        }
    }
}
