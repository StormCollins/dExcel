namespace dExcel;

using ExcelDna.Integration;
using ExcelDna.Registration;
using System.Linq;
using System.Windows.Media.Media3D;
using mnd = MathNet.Numerics.Distributions;
using mnl = MathNet.Numerics.LinearAlgebra;
using mnr = MathNet.Numerics.Random;
using mns = MathNet.Numerics.Statistics;

public static class StatsUtils
{
    [ExcelFunction(
        Name = "d.Stats_Cholesky",
        Description = "Calculates the Cholesky decomposition of a symmetric positive-definite matrix.\n" +
                      "Deprecates the AQS function: 'Chol'",
        Category = "∂Excel: Stats")]
    public static object Cholesky(
        [ExcelArgument(
            Name = "Range",
            Description = "The range containing the NxN (square) matrix.")]
        double[,] range)
    {
        var matrix = mnl.CreateMatrix.DenseOfArray(range);

        if(matrix.RowCount != matrix.ColumnCount)
        {
            return CommonUtils.DExcelErrorMessage("Matrix supplied is not square.");
        }
        else if (matrix.IsSymmetric() == false)
        {
            return CommonUtils.DExcelErrorMessage("Matrix supplied is not symmetric.");
        }
        else
        {
            try
            {
                return matrix.Cholesky().Factor.ToArray();
            }
            catch
            {
                return CommonUtils.DExcelErrorMessage("Matrix supplied is not positive-definite.");
            }
            
        }
        
    }

    [ExcelFunction(
        Name = "d.Stats_CorrelationMatrix",
        Description = "Calculates the Pearson correlation matrix.\n" +
                      "Deprecates the AQS function 'Corr'.",
        Category = "∂Excel: Stats")]
    public static double[,] CorrelationMatrix(
        [ExcelArgument(
            Name = "Range",
            Description = "The range containing the column-wise data.")]
        double[,] range)
    {
        var data = new double[range.GetLength(1)][];
        for (int j = 0; j < range.GetLength(1); j++)
        {
            data[j] = new double[range.GetLength(0)];
            for (int i = 0; i < range.GetLength(0); i++)
            {
                data[j][i] = range[i, j];
            }
        }
        return mns.Correlation.PearsonMatrix(data).ToArray();
    }

    [ExcelFunction(
        Name = "d.Stats_NormalRandomNumbers",
        Description = "Generates a sequence of standard normal random variates.\n" +
                      "Deprecates AQS function: 'Randn'",
        Category = "∂Excel: Stats",
        IsVolatile = true)]
    public static object NormalRandomNumbers(
    [ExcelArgument(
            Name = "Seed",
            Description = "The seed for the random number generator. If left blank, a random seed will be used.")]
        object seed)
    {
        var caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
        var rowCount = caller.RowLast - caller.RowFirst + 1;
        var columnCount = caller.ColumnLast - caller.ColumnFirst + 1;
        var results = new double[rowCount, columnCount];
        var random = new mnr.MersenneTwister();

        if (seed is not ExcelDna.Integration.ExcelMissing)
        {
            try
            {
                random = new mnr.MersenneTwister(Convert.ToInt32(seed));
            }
            catch
            {
                return CommonUtils.DExcelErrorMessage("Invalid seed specified.");
            }
        }
        
        for (int j = 0; j < columnCount; j++)
        {
            for (int i = 0; i < rowCount; i++)
            {
                results[i, j] = mnd.Normal.InvCDF(0.0, 1.0, random.NextDouble());
            }
        }

        return results;
    }
}
