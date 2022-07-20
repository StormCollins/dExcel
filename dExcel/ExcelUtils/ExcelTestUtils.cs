namespace dExcel.ExcelUtils;

using System;
using ExcelDna.Integration;

/// <summary>
/// A collection of utilities for performing basic fuzzy logic on tests/checks in Excel.
/// </summary>
public static class ExcelTestUtils
{
    /// <summary>
    /// The possible outputs from a test/check in Excel.
    /// </summary>
    private enum TestOutputs
    {
        OK, ERROR, WARNING,
    }

    /// <summary>
    /// Checks if two numeric or string values in Excel are the same.
    /// </summary>
    /// <param name="a">Input a</param>
    /// <param name="b">Input b</param>
    /// <returns>'OK' if the values are equal otherwise 'ERROR'.</returns>
    [ExcelFunction(
        Name = "d.Test_Equal",
        Description = "Returns 'OK' if two values are equal, otherwise it returns 'ERROR'.",
        Category = "∂Excel: Test")]
    public static string Equal(
        [ExcelArgument(
            Name = "a",
            Description = "Input a")]
        object a,
        [ExcelArgument(
            Name = "b",
            Description = "Input b")]
        object b)
    {
        if (double.TryParse(a.ToString(), out var x) && double.TryParse(b.ToString(), out var y))
        {
            return Math.Abs(y - x) < 0.00000000001 ? TestOutputs.OK.ToString() : TestOutputs.ERROR.ToString();
        }

        return string.Compare(a.ToString(), b.ToString(), StringComparison.OrdinalIgnoreCase) == 0
            ? TestOutputs.OK.ToString()
            : TestOutputs.ERROR.ToString();
    }

    /// <summary>
    /// Returns 'OK' if input is true, otherwise returns 'ERROR'.
    /// </summary>
    /// <param name="x">x</param>
    /// <returns>Returns 'OK' if input is true, otherwise returns 'ERROR'.</returns>
    [ExcelFunction(
        Name = "d.Test_IsTrue",
        Description = "Returns 'OK' if input is true, otherwise returns 'ERROR'.",
        Category = "∂Excel: Test")]
    public static object IsTrue(
        [ExcelArgument(
            Name = "x",
            Description = "Boolean input.")]
        object x)
        => (bool)x ? TestOutputs.OK.ToString() : TestOutputs.ERROR.ToString();

    /// <summary>
    /// Returns 'OK' if input is false, otherwise returns 'ERROR'.
    /// </summary>
    /// <param name="x">x</param>
    /// <returns>Returns 'OK' if input is false, otherwise returns 'ERROR'.</returns>
    [ExcelFunction(
        Name = "d.Test_IsFalse",
        Description = "Returns 'Okay' if parameter is false, otherwise returns 'Error'.",
        Category = "∂Excel: Test")]
    public static object IsFalse(
        [ExcelArgument(
            Name = "x",
            Description = "Boolean input.")]
        object x)
        => !(bool)x ? TestOutputs.OK.ToString() : TestOutputs.ERROR.ToString();

    /// <summary>
    /// Checks if input 'a' is strictly greater than input 'b'.
    /// </summary>
    /// <param name="a">Input a</param>
    /// <param name="b">Input b</param>
    /// <returns>Returns 'OK' if one value is greater than the other, otherwise it returns 'Error'.</returns>
    [ExcelFunction(
        Name = "d.Test_GreaterThan",
        Description = "Returns 'OK' if input 'a' is strictly greater than input 'b', otherwise 'ERROR'.",
        Category = "∂Excel: Test")]
    public static string GreaterThan(
        [ExcelArgument(
            Name = "a",
            Description = "Input a")]
        double a,
        [ExcelArgument(
            Name = "b",
            Description = "Input b")]
        double b) =>
        (double)a > (double)b ? TestOutputs.OK.ToString() : TestOutputs.ERROR.ToString();

    /// <summary>
    /// Checks if input 'a' is strictly less than input 'b'.
    /// </summary>
    /// <param name="a">Input a</param>
    /// <param name="b">Input b</param>
    /// <returns>Returns 'OK' if input 'a' is strictly less than input 'b', otherwise it returns 'Error'.</returns>
    [ExcelFunction(
        Name = "d.Test_LessThan",
        Description = "Returns 'OK' if input 'a' is strictly less than input 'b', otherwise it returns 'ERROR'.",
        Category = "∂Excel: Test")]
    public static string LessThan(
        [ExcelArgument(
            Name = "a",
            Description = "Input a")]
        double a,
        [ExcelArgument(
            Name = "b",
            Description = "Input b")]
        double b) =>
        (double)a < (double)b ? TestOutputs.OK.ToString() : TestOutputs.ERROR.ToString();

    /// <summary>
    /// Acts like a fuzzy logic 'And' with the following rules.
    ///   'ERROR' and X = 'ERROR'
    ///   'WARNING' and 'OK' = 'WARNING'
    /// i.e. 'ERROR' > 'WARNING' > 'WARNING' it can be seen as checking that there are only 'OK's.
    /// </summary>
    /// <param name="xRange">The input range.</param>
    /// <returns>'ERROR' if the input range contains any 'ERROR's, otherwise 'WARNING' if there are no 'ERROR's,
    /// otherwise 'OK'.</returns>
    [ExcelFunction(
        Name = "d.Test_And",
        Description = "Acts like a fuzzy logic 'And' with the following rules." +
        "\n  'ERROR' and X = 'ERROR'" +
        "\n  'WARNING' and 'OK' = 'WARNING'" +
        "\ni.e. 'ERROR' > 'WARNING' > 'OK' it can be seen as checking that there are only 'OK's.",
        Category = "∂Excel: Test")]
    public static string And(
        [ExcelArgument(
            Name = "Range",
            Description = "Input range.")]
        params object[] xRange)
    {
        var resultArray = new object[xRange.Length];

        for (int i = 0; i < xRange.Length; i++)
        {
            if (xRange[i].GetType() == typeof(object[]))
            {
                resultArray[i] = And1d((object[])xRange[i]);
            }
            else if (xRange[i].GetType() == typeof(object[,]))
            {
                resultArray[i] = And2d((object[,])xRange[i]);
            }
            else
            {
                resultArray[i] = xRange[i].ToString() ?? string.Empty;
            }
        }

        return And1d(resultArray);

        string And1d(object[] x)
        {
            var result = TestOutputs.OK.ToString();
            for (int i = 0; i < x.GetLength(0); i++)
            {
                if (string.Compare(x[i].ToString(), TestOutputs.ERROR.ToString(), StringComparison.OrdinalIgnoreCase) == 0)
                {
                    return TestOutputs.ERROR.ToString();
                }
                
                if (string.Compare(x[i].ToString(), TestOutputs.WARNING.ToString(), StringComparison.OrdinalIgnoreCase) == 0)
                {
                    result = TestOutputs.WARNING.ToString();
                }
                else
                {
                    result = result == TestOutputs.WARNING.ToString() ? TestOutputs.WARNING.ToString() : TestOutputs.OK.ToString();
                }
            }

            return result;
        }

        string And2d(object[,] x)
        {
            var result = TestOutputs.OK.ToString();
            for (int i = 0; i < x.GetLength(0); i++)
            {
                for (int j = 0; j < x.GetLength(1); j++)
                {
                    if (string.Compare(x[i, j].ToString(), TestOutputs.ERROR.ToString(), StringComparison.OrdinalIgnoreCase) == 0)
                    {
                        return TestOutputs.ERROR.ToString();
                    }
                    
                    if (string.Compare(x[i, j].ToString(), TestOutputs.WARNING.ToString(), StringComparison.OrdinalIgnoreCase) == 0)
                    {
                        result = TestOutputs.WARNING.ToString();
                    }
                    else
                    {
                        result = result == TestOutputs.WARNING.ToString() ? TestOutputs.WARNING.ToString() : TestOutputs.OK.ToString();
                    }
                }
            }

            return result;
        }
    }

    /// <summary>
    /// Acts as a fuzzy logic 'Not'.
    /// </summary>
    /// <param name="x"></param>
    /// <returns>Returns 'OK' if input is 'ERROR', 'ERROR' if 'OK' and 'WARNING' otherwise.</returns>
    [ExcelFunction(
        Name = "d.Test_Not",
        Description = "Returns 'OK' if input is 'ERROR', 'ERROR' if 'OK' and 'WARNING' otherwise.",
        Category = "∂Excel: Test")]
    public static object Not(
        [ExcelArgument(
            Name = "Range",
            Description = "Input range.")]
        object[,] x)
    {
        var output = new object[x.GetLength(0), x.GetLength(1)];

        for (int i = 0; i < x.GetLength(0); i++)
        {
            for (int j = 0; j < x.GetLength(1); j++)
            {
                if (string.Compare(x[i, j].ToString(), TestOutputs.OK.ToString(), StringComparison.OrdinalIgnoreCase) == 0)
                {
                    output[i, j] = TestOutputs.ERROR.ToString();
                }
                else if (string.Compare(x[i, j].ToString(), TestOutputs.ERROR.ToString(), StringComparison.OrdinalIgnoreCase) == 0)
                {
                    output[i, j] = TestOutputs.OK.ToString();
                }
                else
                {
                    output[i, j] = TestOutputs.WARNING.ToString();
                }
            }
        }
        
        return output;
    }

    /// <summary>
    /// Acts like a fuzzy logic 'Or' with the following rules.
    ///   'OK' or X = 'OK'
    ///   'WARNING' or 'ERROR' = 'WARNING'
    /// i.e. 'OK' > 'WARNING' > 'ERROR' it can be seen as finding ANY 'OK's.
    /// </summary>
    /// <param name="xRange">The input range.</param>
    /// <returns>'OK' if the input range contains any 'OK's, otherwise 'WARNING' if there are no 'OK's, otherwise 'ERROR'.</returns>
    [ExcelFunction(
        Name = "d.Test_Or",
        Description = "Acts like a fuzzy logic 'Or' with the following rules." +
        "\n  'OK' or X = 'OK'" +
        "\n  'WARNING' or 'ERROR' = 'WARNING'" +
        "\ni.e. 'OK' > 'WARNING' > 'ERROR' it can be seen as finding ANY 'OK's.",
        Category = "∂Excel: Test")]
    public static string Or(
        [ExcelArgument(
            Name = "Range",
            Description = "Input range.")]
        params object[] xRange)
    {
        var resultArray = new object[xRange.Length];

        for (int i = 0; i < xRange.Length; i++)
        {
            if (xRange[i].GetType() == typeof(object[]))
            {
                resultArray[i] = Or1d((object[])xRange[i]);
            }
            else if (xRange[i].GetType() == typeof(object[,]))
            {
                resultArray[i] = Or2d((object[,])xRange[i]);
            }
            else
            {
                resultArray[i] = xRange[i].ToString() ?? string.Empty;
            }
        }

        return Or1d(resultArray);

        string Or1d(object[] x)
        {
            var result = TestOutputs.OK.ToString();
            for (int i = 0; i < x.GetLength(0); i++)
            {
                if (string.Compare(x[i].ToString(), TestOutputs.OK.ToString(), StringComparison.OrdinalIgnoreCase) == 0)
                {
                    return TestOutputs.OK.ToString();
                }
                
                if (string.Compare(x[i].ToString(), TestOutputs.WARNING.ToString(), StringComparison.OrdinalIgnoreCase) == 0)
                {
                    result = TestOutputs.WARNING.ToString();
                }
                else
                {
                    result = result == TestOutputs.WARNING.ToString() ? TestOutputs.WARNING.ToString() : TestOutputs.ERROR.ToString();
                }
            }
            
            return result;
        }

        string Or2d(object[,] x)
        {
            var result = TestOutputs.OK.ToString();
            for (int i = 0; i < x.GetLength(0); i++)
            {
                for (int j = 0; j < x.GetLength(1); j++)
                {
                    if (string.Compare(x[i, j].ToString(), TestOutputs.OK.ToString(), StringComparison.OrdinalIgnoreCase) == 0)
                    {
                        return TestOutputs.OK.ToString();
                    }
                    
                    if (string.Compare(x[i, j].ToString(), TestOutputs.WARNING.ToString(), StringComparison.OrdinalIgnoreCase) == 0)
                    {
                        result = TestOutputs.WARNING.ToString();
                    }
                    else
                    {
                        result = result == TestOutputs.WARNING.ToString() ? TestOutputs.WARNING.ToString() : TestOutputs.ERROR.ToString();
                    }
                }
            }
            
            return result;
        }
    }
}
