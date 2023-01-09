namespace dExcel.Mathematics;

using ExcelDna.Integration;
using Utilities;
using mnd = MathNet.Numerics.Distributions;
using mnl = MathNet.Numerics.LinearAlgebra;
using mnr = MathNet.Numerics.Random;
using mns = MathNet.Numerics.Statistics;

public static class StatsUtils
{
    /// <summary>
    /// Calculates the Cholesky decomposition of a symmetric positive-definite matrix.
    /// Note this returns an upper-triangular matrix.
    /// </summary>
    /// <param name="range">The range containing the NxN (square) matrix.</param>
    /// <returns>The Cholesky decomposition of a matrix if valid.</returns>
    [ExcelFunction(
        Name = "d.Stats_Cholesky",
        Description = "Calculates the Cholesky decomposition of a symmetric positive-definite matrix.\n" +
                      "Note this returns an upper-triangular matrix.\n" +
                      "Deprecates the AQS function: 'Chol'",
        Category = "∂Excel: Stats")]
    public static object Cholesky(
        [ExcelArgument(
            Name = "Range",
            Description = "The range containing the NxN (square) matrix.")]
        double[,] range)
    {
        mnl.Matrix<double>? matrix = mnl.CreateMatrix.DenseOfArray(range);

        if(matrix.RowCount != matrix.ColumnCount)
        {
            return CommonUtils.DExcelErrorMessage("Matrix is not square.");
        }

        if (matrix.IsSymmetric() == false)
        {
            return CommonUtils.DExcelErrorMessage("Matrix is not symmetric.");
        }

        try
        {
            return matrix.Cholesky().Factor.Transpose().ToArray();
        }
        catch
        {
            return CommonUtils.DExcelErrorMessage("Matrix is not positive-definite.");
        }
    }

    /// <summary>
    /// Calculates the Pearson correlation matrix.
    /// </summary>
    /// <param name="range">The range containing the column-wise data.</param>
    /// <returns>The Pearson correlation matrix.</returns>
    [ExcelFunction(
        Name = "d.Stats_CorrelationMatrix",
        Description = "Calculates the Pearson correlation matrix.\n" +
                      "Deprecates the AQS function 'Corr'.",
        Category = "∂Excel: Stats")]
    public static object CorrelationMatrix(
        [ExcelArgument(
            Name = "Range",
            Description = "The range containing the column-wise data.")]
        double[,] range)
    {
        double[][] data = new double[range.GetLength(1)][];
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

    /// <summary>
    /// Generates a sequence of standard normal random variates.
    ///
    /// The total number of elements returned is given either by the size of the region the user has selected in Excel or
    /// the optional parameters "rowCount" and "columnCount".
    /// </summary>
    /// <param name="seed">Seed</param>
    /// <param name="rowCount">The number of rows to output.</param>
    /// <param name="columnCount">The number of columns to output.</param>
    /// <returns>A region of standard normal random variates.</returns>
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
        int seed,
    [ExcelArgument(
        Name = "(Optional)Row Count",
        Description = "The number of rows of random numbers to output.")]
        int rowCount = 1,
    [ExcelArgument(
        Name = "(Optional)Column Count",
        Description = "The number of columns of random numbers to output.")]
        int columnCount = 1)
    {
        if (ExcelDnaUtil.Application is not null)
        {
            ExcelReference? caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
            rowCount = caller.RowLast - caller.RowFirst + 1;
            columnCount = caller.ColumnLast - caller.ColumnFirst + 1;
        }

        double[,] results = new double[rowCount, columnCount];
        mnr.MersenneTwister random = new();
        random = new mnr.MersenneTwister(Convert.ToInt32(seed));
        
        for (int j = 0; j < columnCount; j++)
        {
            for (int i = 0; i < rowCount; i++)
            {
                results[i, j] = mnd.Normal.InvCDF(0.0, 1.0, random.NextDouble());
            }
        }

        return results;
    }

    // public static object CorrelatedNormalRandomNumbers(int rowCount, double[,] correlationMatrix)
    // {
    //      
    // }
}
