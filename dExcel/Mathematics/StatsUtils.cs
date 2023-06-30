using System.Windows.Automation;
using ExcelDna.Integration;
using dExcel.Utilities;
using mnd = MathNet.Numerics.Distributions;
using mnl = MathNet.Numerics.LinearAlgebra;
using mnr = MathNet.Numerics.Random;
using mns = MathNet.Numerics.Statistics;
using QL = QuantLib;

namespace dExcel.Mathematics;

/// <summary>
/// A collection of statistical utility functions.
/// </summary>
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
        Description = 
            "Calculates the Cholesky decomposition of a symmetric positive-definite matrix.\n" +
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

        for (int i = 0; i < range.GetLength(0); i++)
        {
            for (int j = 0; j < range.GetLength(1); j++)
            {
                if (i == j && Math.Abs(range[i, j] - 1.0) > 0.0001)
                {
                    return CommonUtils.DExcelErrorMessage("Diagonal elements of the correlation matrix must be 1.");
                }
            }
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
    /// The total number of elements returned is given either by the size of the region the user has selected in Excel
    /// or the optional parameters "rowCount" and "columnCount".
    /// </summary>
    /// <param name="seed">Seed</param>
    /// <param name="rowCount">The number of rows to output.</param>
    /// <param name="columnCount">The number of columns to output.</param>
    /// <returns>A region of standard normal random variates.</returns>
    [ExcelFunction(
        Name = "d.Stats_NormalRandomNumbers",
        Description = 
            "Generates a sequence of standard normal random variates.\n" +
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
        int rowCount = 0,
    [ExcelArgument(
        Name = "(Optional)Column Count",
        Description = "The number of columns of random numbers to output.")]
        int columnCount = 0)
    {
        if (ExcelDnaUtil.Application is not null && rowCount == 0 && columnCount == 0)
        {
            ExcelReference? caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
            if (caller != null)
            {
                rowCount = caller.RowLast - caller.RowFirst + 1;
                columnCount = caller.ColumnLast - caller.ColumnFirst + 1;
            }
        }

        double[,] results = new double[rowCount, columnCount];
        mnr.MersenneTwister random = new(Convert.ToInt32(seed));
        
        for (int j = 0; j < columnCount; j++)
        {
            for (int i = 0; i < rowCount; i++)
            {
                results[i, j] = mnd.Normal.InvCDF(0.0, 1.0, random.NextDouble());
            }
        }
        
        return results;
    }

    /// <summary>
    /// Returns a set of correlated, normal random numbers.
    /// </summary>
    /// <param name="seed">Seed for the random number generator.</param>
    /// <param name="correlatedSetCount">The number of sets of correlated random numbers to generate.
    /// e.g., If this is 'm' and the size of the correlation matrix is 'n x n' then the number of random numbers
    /// generated will be 'mn'.</param>
    /// <param name="correlationMatrixRange">Correlation matrix of the random numbers.</param>
    /// <returns>A set of correlated, normal random numbers.</returns>
    [ExcelFunction(
        Name = "d.Stats_CorrelatedNormalRandomNumbers",
        Description = "Returns a set of correlated, normal random numbers.",
        Category = "∂Excel: Stats")]
    public static object CorrelatedNormalRandomNumbers(
        [ExcelArgument(
            Name = "Seed",
            Description = "Seed for the random number generator.")]
        int seed,
        [ExcelArgument(
            Name = "Correlated Set Count",
            Description = 
                "The number of sets of correlated random numbers to generate.\n" +
                "e.g., If this is 'm' and the size of the correlation matrix is 'n x n' then the number of random" +
                "numbers generated will be 'mn'.")]
        int correlatedSetCount,
        [ExcelArgument(
            Name = "Correlation Matrix", 
            Description = "Correlation matrix of the random numbers.")]
        double[,] correlationMatrixRange)
    {
        mnl.Matrix<double> randomNumbers = 
            mnl.CreateMatrix.DenseOfArray(
                (double[,])NormalRandomNumbers(seed, correlatedSetCount, correlationMatrixRange.GetLength(0)));
        
        object choleskyResults =  Cholesky(correlationMatrixRange);
        if (choleskyResults is string errorMessage)
        {
            return errorMessage;
        }
        
        mnl.Matrix<double> choleskyMatrix = mnl.CreateMatrix.DenseOfArray((double[,])choleskyResults);
        mnl.Matrix<double> results = randomNumbers * choleskyMatrix;
        return results.ToArray();
    }

    /// <summary>
    /// Creates a GBM object which can be queried for GBM paths.
    /// </summary>
    /// <param name="handle">The 'handle' or name used to refer to the object in memory.</param>
    /// <param name="initialValue">The initial value to be simulated.</param>
    /// <param name="drift">The drift of the GBM process.</param>
    /// <param name="standardDeviation">The standard deviation of the GBM process.</param>
    /// <param name="maturityInYears"></param>
    /// <param name="numberOfTimeSteps"></param>
    /// <param name="seed">The seed of the random number generator.</param>
    /// <returns>A handle to a GBM object.</returns>
    [ExcelFunction(
        Name = "d.Stats_GBM_Create",
        Description = "Creates a GBM object which can be queried for GBM paths.",
        Category = "∂Excel: Stats")]
    public static object GbmCreate(
        [ExcelArgument(
            Name = "Handle",
            Description = "The 'handle' or name used to refer to the object in memory.")]
        string handle,
        [ExcelArgument(Name = "Initial Value", Description = "The initial value to be simulated.")]
        double initialValue,
        [ExcelArgument(Name = "Drift", Description = "The drift of the GBM process.")]
        double drift,
        [ExcelArgument(Name = "Standard Deviation", Description = "The standard deviation of the GBM process.")]
        double standardDeviation,
        [ExcelArgument(Name = "Maturity in Years", Description = "The maturity of the simulation in years.")]
        double maturityInYears,
        [ExcelArgument(
            Name = "Number of Time Steps", 
            Description = "Number of time steps in the Monte Carlo simulation.")]
        int numberOfTimeSteps,
        [ExcelArgument(Name = "Seed", Description = "The seed of the random number generator.")]
        int seed)
    {
        if (ExcelDnaUtil.IsInFunctionWizard())
        {
            return CommonUtils.InFunctionWizard();
        }

        QL.UniformRandomGenerator uniformRandomGenerator = new(seed);
        QL.UniformRandomSequenceGenerator uniformSequenceGenerator = new((uint)numberOfTimeSteps, uniformRandomGenerator);
        QL.GaussianRandomSequenceGenerator gaussianSequenceRandomGenerator = new(uniformSequenceGenerator);
        QL.GeometricBrownianMotionProcess gbmProcess = new(initialValue, drift, standardDeviation);
        QL.GaussianPathGenerator gaussianPathGenerator =
            new(gbmProcess, maturityInYears, (uint)numberOfTimeSteps, gaussianSequenceRandomGenerator, false);

        DataObjectController instance = DataObjectController.Instance;
        return instance.Add(handle, gaussianPathGenerator);
    }

    /// <summary>
    /// Gets the paths from a GBM object that has been created.
    /// </summary>
    /// <param name="handle">The 'handle' or name used to refer to the object in memory.</param>
    /// <param name="orientation">The orientation as to how to output the paths either as 'ROWS' or 'COLUMNS'.
    /// Default = 'COLUMNS'</param>
    /// <returns>The numeric path values from the GBM Monte Carlo simulation.</returns>
    [ExcelFunction(
        Name = "d.Stats_GBM_GetPaths",
        Description = "Gets the paths from a GBM object that has been created.",
        Category = "∂Excel: Stats")]
    public static object GbmGetPaths(
        [ExcelArgument(
            Name = "Handle",
            Description = "The 'handle' or name used to refer to the object in memory.")]
        string handle,
        [ExcelArgument(Name = "Number of Paths", Description = "Number of paths in the Monte Carlo simulation.")]
        int numberOfPaths,
        [ExcelArgument(
            Name = "(Optional)Orientation",
            Description = "The orientation as to how to output the paths either as 'ROWS' or 'COLUMNS'.\n" +
                          "Default = 'COLUMNS'")]
        string orientation = "ROWS")
    {
        DataObjectController instance = DataObjectController.Instance;
        QL.GaussianPathGenerator pathGenerator = (QL.GaussianPathGenerator)instance.GetDataObject(handle);

        if (orientation.IgnoreCaseEquals("ROWS"))
        {
            object[,] output = new object[pathGenerator.size(), pathGenerator.timeGrid().size() + 1];
            for (int i = 0; i < numberOfPaths; i++)
            {
                QL.Path? path = pathGenerator.next().value();
                for (int j = 0; j < pathGenerator.timeGrid().size(); j++)
                {
                    output[i, j] = path.value((uint)j);
                }
            }

            return output;
        }

        if (orientation.IgnoreCaseEquals("ROWS"))
        {
            object[,] output = new object[pathGenerator.timeGrid().size() + 1, pathGenerator.size()];
            for (int i = 0; i < numberOfPaths; i++)
            {
                QL.Path? path = pathGenerator.next().value();
                for (int j = 0; j < pathGenerator.timeGrid().size(); j++)
                {
                    output[j, i] = path.value((uint)j);
                }
            }

            return output;
        }

        return CommonUtils.DExcelErrorMessage($"Invalid orientation: '{orientation}'");
    }
}
