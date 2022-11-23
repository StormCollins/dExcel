namespace dExcel.Mathematics;

using System;
using System.Linq;
using ExcelDna.Integration;
using MathNet.Numerics;
using mni = MathNet.Numerics.Interpolation;

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
        int rowCount = xy.GetLength(0) - 1;
        int colCount = xy.GetLength(1) - 1;
        double[] rowValues = new double[rowCount];
        double[] colValues = new double[colCount];

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
            return $"{CommonUtils.DExcelErrorPrefix} Extrapolation not supported.";
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
        Name = "d.Math_Interpolate",
        Description = "Performs linear, exponential, or flat interpolation on a range for a single point.\n" +
                      "Deprecates AQS function: 'dt_interp'",
        Category = "∂Excel: Mathematics")]
    public static object Interpolate(
        [ExcelArgument(Name = "X-column", Description = "Independent variable.")]
        object[,] xCol,
        [ExcelArgument(Name = "Y-column", Description = "Dependent variable.")]
        object[,] yCol,
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

        if ((xCol.GetLength(0) == yCol.GetLength(0)) && (xCol.GetLength(1)==1) && (yCol.GetLength(1) == 1))
        {
            int rowCount = Math.Max(xCol.GetLength(0), yCol.GetLength(0));
            double[] x = new double[rowCount];
            double[] y = new double[rowCount];
            for (int i = 0; i < rowCount; i++)
            {
                x[i] = (double)xCol[i, 0];
                y[i] = (double)yCol[i, 0];
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
        }
        
        return CommonUtils.DExcelErrorMessage("Row dimensions do not match or there is more than one column in x or y.");
    }
}
