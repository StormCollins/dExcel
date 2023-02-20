namespace dExcelTests.Mathematics;

using dExcel.Mathematics;
using dExcel.Utilities;
using mnd = MathNet.Numerics.Distributions;
using mnla = MathNet.Numerics.LinearAlgebra;
using mnr = MathNet.Numerics.Random;
using mns = MathNet.Numerics.Statistics;
using NUnit.Framework;

[TestFixture]
public class StatsUtilsTests
{
   [Test]
   public void CholeskyTest()
   {
      double[,] correlationMatrix =
      {
         {1.0, 0.5},
         {0.5, 1.0}
      };
      
      double[,] expected =
      {
         {1.0, 0.5},
         {0.0, Math.Sqrt(1 - 0.5 * 0.5)}
      };
      
      double[,] actual = (double[,])StatsUtils.Cholesky(correlationMatrix);
      
      Assert.AreEqual(expected, actual);
   }

   [Test]
   public void NonSquareMatrixCholeskyTest()
   {
      double[,] correlationMatrix =
      {
         {1.0, 0.5, 0.0},
         {0.5, 1.0, 0.0}
      };
      
      string? actual = StatsUtils.Cholesky(correlationMatrix).ToString();
      string expected = CommonUtils.DExcelErrorMessage("Matrix is not square.");
      Assert.AreEqual(expected, actual);
   }

   [Test]
   public void NonSymmetricMatrixCholeskyTest()
   {
      double[,] correlationMatrix =
      {
         {1.0, 0.5},
         {0.0, 1.0}
      };
      
      string? actual = StatsUtils.Cholesky(correlationMatrix).ToString();
      string expected = CommonUtils.DExcelErrorMessage("Matrix is not symmetric.");
      Assert.AreEqual(expected, actual);
   }

   [Test]
   public void NonPositiveDefiniteMatrixCholeskyTest()
   {
      double[,] correlationMatrix =
      {
         {1.0, 0.0},
         {0.0, -1.0}
      };
      
      string? actual = StatsUtils.Cholesky(correlationMatrix).ToString();
      string expected = CommonUtils.DExcelErrorMessage("Diagonal elements of the correlation matrix must be 1.");
      Assert.AreEqual(expected, actual);
   }

   [Test]
   public void CorrelationMatrixTest()
   {
      double[,] correlationMatrix =
      {
         {1.0, 0.5},
         {0.5, 1.0}
      };

      mnla.Matrix<double> cholesky = mnla.CreateMatrix.DenseOfArray((double[,])StatsUtils.Cholesky(correlationMatrix));

      const int totalRandomNumbers = 100_000;
      double[]? uniformRandomNumbers = mnr.SystemRandomSource.Doubles(totalRandomNumbers, 999);
      double[,] uncorrelatedRandomNumbers = new double[totalRandomNumbers/2, 2];
      for (int i = 0; i < totalRandomNumbers/2; i++)
      {
         for (int j = 0; j < 2; j++)
         {
            uncorrelatedRandomNumbers[i, j] = uniformRandomNumbers[(totalRandomNumbers/2) * j + i];
         }
      }

      double[,] correlatedRandomNumbers =
         (mnla.CreateMatrix.DenseOfArray(uncorrelatedRandomNumbers) * cholesky).ToArray();
      double[,] actual = (double[,])StatsUtils.CorrelationMatrix(correlatedRandomNumbers);
      double[,] expected = 
      {
         {1.0, 0.5},
         {0.5, 1.0}
      };

      for (int i = 0; i < actual.GetLength(0); i++)
      {
         for (int j = 0; j < actual.GetLength(1); j++)
         {
            Assert.AreEqual(expected[i, j], actual[i, j], 0.005);   
         }
      }
   }

   [Test]
   public void NormalRandomNumbersTest()
   {
      double[,] actual = (double[,])StatsUtils.NormalRandomNumbers(999, 1, 1);
      double[,] expected = {{ 0.85392973639129077 }};
      Assert.AreEqual(expected, actual);
   }

   [Test]
   public void CorrelatedRandomNumbersTest()
   {
      double[,] correlationMatrix = {{ 1.0, 0.5 }, { 0.5, 1.0 }};
      double[,] correlatedRandomNumbers = (double[,])StatsUtils.CorrelatedNormalRandomNumbers(999, 100_000, correlationMatrix);
      double[] randomNumberSet1 = new double[correlatedRandomNumbers.GetLength(0)];
      double[] randomNumberSet2 = new double[correlatedRandomNumbers.GetLength(0)];
      for (int i = 0; i < correlatedRandomNumbers.GetLength(0); i++)
      {
         randomNumberSet1[i] = correlatedRandomNumbers[i, 0];     
         randomNumberSet2[i] = correlatedRandomNumbers[i, 1];
      } 
      
      double actualCorrelation = mns.Correlation.Pearson(randomNumberSet1, randomNumberSet2); 
      const double expectedCorrelation = 0.5;
      Assert.AreEqual(expectedCorrelation, actualCorrelation, 0.01);
   }

   [Test]
   public void NonSquareMatrixForCorrelatedNormalRandomNumbersTest()
   {
      double[,] correlationMatrix = {{ 1.0, 0.5, 0.1 }, { 0.1, 1.0, 1.0 }};
      string actual = StatsUtils.CorrelatedNormalRandomNumbers(999, 100_000, correlationMatrix).ToString();
      string expected = CommonUtils.DExcelErrorMessage("Matrix is not square.");
      Assert.AreEqual(expected, actual);
   }

   [Test]
   public void DiagonalCorrelationNotOneForCorrelatedNormalRandomNumbersTest()
   {
      double[,] correlationMatrix = {{ 1.0, 0.5 }, { 0.5, 2.0 }};
      string actual = StatsUtils.CorrelatedNormalRandomNumbers(999, 100_000, correlationMatrix).ToString();
      string expected = CommonUtils.DExcelErrorMessage("Diagonal elements of the correlation matrix must be 1.");
      Assert.AreEqual(expected, actual);
   }
}
