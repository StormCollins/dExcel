namespace dExcel;

using ExcelDna.Integration;
using mnd = MathNet.Numerics.Distributions;
using mnl = MathNet.Numerics.LinearAlgebra;
using mnr = MathNet.Numerics.Random;
using mns = MathNet.Numerics.Statistics;

public static class StatsUtils
{
    [ExcelFunction(
        Name = "d.Chol",
        Description = "Calculates the Cholesky decomposition of a matrix.\n" +
                      "Deprecates AQS function: 'Chol'",
        Category = "∂Excel: Stats")]
    public static double[,] CholeskyDecomposition(
        [ExcelArgument(
            Name = "Range",
            Description = "The range containing the nxn matrix.")]
        double[,] range)
    {
        var matrix = mnl.CreateMatrix.DenseOfArray(range);
        return matrix.Cholesky().Factor.Transpose().ToArray();
    }

    [ExcelFunction(
        Name = "d.Corr",
        Description = "Calculates the Pearson correlation matrix." +
                      "\nDeprecates the AQS function 'corr'.",
        Category = "∂Excel: Stats")]
    public static double[,] Correlation(
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
        Name = "d.Randn",
        Description = "Generates standard random normal variates.\n" +
                      "Deprecates AQS function: 'Randn'",
        Category = "∂Excel: Stats")]
    public static double[,] Randn(
    [ExcelArgument(
            Name = "Seed",
            Description = "The seed for the random number generator.")]
        int seed)
    {
        var caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
        var rowCount = caller.RowLast - caller.RowFirst + 1;
        var columnCount = caller.ColumnLast - caller.ColumnFirst + 1;
        var results = new double[rowCount, columnCount];
        var random = new mnr.MersenneTwister(seed);
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
