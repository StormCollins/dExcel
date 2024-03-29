﻿using dExcel.Utilities;
using ExcelDna.Integration;

namespace dExcel.ExcelUtils;

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
        Ok, Error, Warning
    }

    /// <summary>
    /// Values that are (semantically) equivalent to ERROR in Excel.
    /// </summary>
    private static readonly List<string?> ExcelErrorValues = new()
    {
        ExcelError.ExcelErrorDiv0.ToString(),
        ExcelError.ExcelErrorName.ToString(),
        ExcelError.ExcelErrorNull.ToString(),
        ExcelError.ExcelErrorNum.ToString(),
        ExcelError.ExcelErrorNA.ToString(),
        ExcelError.ExcelErrorRef.ToString(),
        ExcelError.ExcelErrorValue.ToString(),
    };

    /// <summary>
    /// Checks that input from Excel is valid - that it's not one of e.g, #NA, #VALUE etc.
    /// </summary>
    /// <param name="x">The input.</param>
    /// <returns>True if the input is valid, false otherwise.</returns>
    private static bool AreInputsValid(object x) => 
        ExcelErrorValues.All(excelErrorValue => !x.ToString().IgnoreCaseEquals(excelErrorValue));

    /// <summary>
    /// Checks that inputs in the form of an array from Excel aren't invalid e.g., #NA, #VALUE etc.
    /// </summary>
    /// <param name="x">The input array.</param>
    /// <returns>True if the inputs are valid, false otherwise.</returns>
    private static bool AreInputsValid(object[] x)
    {
        foreach (object t in x)
        {
            foreach (string? excelErrorValue in ExcelErrorValues)
            {
                if (t.ToString().IgnoreCaseEquals(excelErrorValue))
                {
                    return false;
                }
            }
        }

        return true;
    }

    /// <summary>
    /// Checks that inputs in the form of an array from Excel aren't invalid e.g. #NA, #VALUE etc.
    /// </summary>
    /// <param name="x">The input array.</param>
    /// <returns>True if the inputs are valid, false otherwise.</returns>
    private static bool AreInputsValid(object[,] x)
    {
        for (int i = 0; i < x.GetLength(0); i++)
        {
            for (int j = 0; j < x.GetLength(1); j++)
            {
                foreach (string? excelErrorValue in ExcelErrorValues)
                {
                    if (x[i, j].ToString().IgnoreCaseEquals(excelErrorValue))
                    {
                        return false;
                    }
                }
            }
        }

        return true;
    }

    /// <summary>
    /// Checks if two numeric or string values in Excel are the same.
    /// </summary>
    /// <param name="a">Input a</param>
    /// <param name="b">Input b</param>
    /// <param name="tolerance">The tolerance for testing numeric equality e.g., default = 0.0000001.</param>
    /// <param name="useWarning">Set to 'True' to use 'Warning' as the output instead of 'Error' in the case of inequality.</param>
    /// <returns>'OK' if the values are equal otherwise 'ERROR'.</returns>
    [ExcelFunction(
        Name = "d.TestUtils_Equal",
        Description = "Returns 'OK' if two values are equal, otherwise it returns 'ERROR'.",
        Category = "∂Excel: Test")]
    public static string Equal(
        [ExcelArgument(Name = "a", Description = "Input a")]
        object a,
        [ExcelArgument(Name = "b", Description = "Input b")]
        object b,
        [ExcelArgument(
            Name = "(Optional)Tolerance",
            Description = "The threshold used for the calculating numeric equality.\n" +
                          "Default = 0.0000001")]
        double tolerance = 0.0000001,
        [ExcelArgument(
            Name = "(Optional)Use Warning",
            Description = "Set to 'True' to use 'Warning' as the output instead of 'Error' in the case of inequality.\n" +
                          "Default = false")]
        bool useWarning = false)
    {
        if (!AreInputsValid(a) || !AreInputsValid(b))
        {
            return TestOutputs.Error.ToString();
        }

        if (double.TryParse(a.ToString(), out double x) && double.TryParse(b.ToString(), out var y))
        {
            return Math.Abs(y - x) < tolerance ? TestOutputs.Ok.ToString() : TestOutputs.Error.ToString();
        }

        return a.ToString().IgnoreCaseEquals(b)
            ? TestOutputs.Ok.ToString()
            : useWarning 
                ? TestOutputs.Warning.ToString() 
                : TestOutputs.Error.ToString();
    }

    /// <summary>
    /// Returns 'OK' if input is true, otherwise returns 'ERROR'.
    /// </summary>
    /// <param name="x">x</param>
    /// <returns>Returns 'OK' if input is true, otherwise returns 'ERROR'.</returns>
    [ExcelFunction(
        Name = "d.TestUtils_IsTrue",
        Description = "Returns 'OK' if input is true, otherwise returns 'ERROR'.",
        Category = "∂Excel: Test")]
    public static object IsTrue(
        [ExcelArgument(Name = "x", Description = "Boolean input.")]
        object x)
    {
        if (!AreInputsValid(x))
        {
            return TestOutputs.Error.ToString();
        }

        return (bool)x ? TestOutputs.Ok.ToString() : TestOutputs.Error.ToString();
    }

    /// <summary>
    /// Returns 'OK' if input is false, otherwise returns 'ERROR'.
    /// </summary>
    /// <param name="x">x</param>
    /// <returns>Returns 'OK' if input is false, otherwise returns 'ERROR'.</returns>
    [ExcelFunction(
        Name = "d.TestUtils_IsFalse",
        Description = "Returns 'Okay' if parameter is false, otherwise returns 'Error'.",
        Category = "∂Excel: Test")]
    public static object IsFalse([ExcelArgument(Name = "x", Description = "Boolean input.")]object x)
    {
        if (!AreInputsValid(x))
        {
            return TestOutputs.Error.ToString();
        }

        return !(bool)x ? TestOutputs.Ok.ToString() : TestOutputs.Error.ToString();
    }

    /// <summary>
    /// Checks if input 'a' is strictly greater than input 'b'.
    /// </summary>
    /// <param name="a">Input a</param>
    /// <param name="b">Input b</param>
    /// <returns>Returns 'OK' if one value is greater than the other, otherwise it returns 'Error'.</returns>
    [ExcelFunction(
        Name = "d.TestUtils_GreaterThan",
        Description = "Returns 'OK' if input 'a' is strictly greater than input 'b', otherwise 'ERROR'.",
        Category = "∂Excel: Test")]
    public static string GreaterThan(
        [ExcelArgument(Name = "a", Description = "Input a")]
        object a,
        [ExcelArgument(Name = "b", Description = "Input b")]
        object b)
    {
        if (!AreInputsValid(a) || !AreInputsValid(b))
        {
            return TestOutputs.Error.ToString();
        }

        return (double)a > (double)b? TestOutputs.Ok.ToString() : TestOutputs.Error.ToString();
    }

    /// <summary>
    /// Checks if input 'a' is strictly less than input 'b'.
    /// </summary>
    /// <param name="a">Input a</param>
    /// <param name="b">Input b</param>
    /// <returns>Returns 'OK' if input 'a' is strictly less than input 'b', otherwise it returns 'Error'.</returns>
    [ExcelFunction(
        Name = "d.TestUtils_LessThan",
        Description = "Returns 'OK' if input 'a' is strictly less than input 'b', otherwise it returns 'ERROR'.",
        Category = "∂Excel: Test")]
    public static string LessThan([ExcelArgument(Name = "a", Description = "Input a")]
        object a,
        [ExcelArgument(Name = "b", Description = "Input b")]
        object b)
    {
        if (!AreInputsValid(a) || !AreInputsValid(b))
        {
            return TestOutputs.Error.ToString();
        }

        return (double)a < (double)b ? TestOutputs.Ok.ToString() : TestOutputs.Error.ToString();
    }

    /// <summary>
    /// Acts like a fuzzy logic 'And' with the following rules.
    ///   'ERROR' and X = 'ERROR'
    ///   'WARNING' and 'OK' = 'WARNING'
    /// i.e. 'ERROR' > 'WARNING' > 'OK' it can be seen as checking that there are only 'OK's.
    /// </summary>
    /// <param name="xRange">The input range.</param>
    /// <returns>'ERROR' if the input range contains any 'ERROR's, otherwise 'WARNING' if there are no 'ERROR's,
    /// otherwise 'OK'.</returns>
    [ExcelFunction(
        Name = "d.TestUtils_And",
        Description = "Acts like a fuzzy logic 'And' with the following rules." +
        "\n  'ERROR' and X = 'ERROR'" +
        "\n  'WARNING' and 'OK' = 'WARNING'" +
        "\ni.e. 'ERROR' > 'WARNING' > 'OK' it can be seen as checking that there are only 'OK's.",
        Category = "∂Excel: Test")]
    public static string And(
        [ExcelArgument(Name = "Range", Description = "Input range.")]
        params object[] xRange)
    {
        object[] resultArray = new object[xRange.Length];

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
            if (!AreInputsValid(x))
            {
                return TestOutputs.Error.ToString();
            }

            string result = TestOutputs.Ok.ToString();
            for (int i = 0; i < x.GetLength(0); i++)
            {
                if (x[i].ToString().IgnoreCaseEquals(TestOutputs.Error))
                {
                    return TestOutputs.Error.ToString();
                }
                
                if (x[i].ToString().IgnoreCaseEquals(TestOutputs.Warning))
                {
                    result = TestOutputs.Warning.ToString();
                }
                else
                {
                    result = result == TestOutputs.Warning.ToString()
                        ? TestOutputs.Warning.ToString()
                        : TestOutputs.Ok.ToString();
                }
            }

            return result;
        }

        string And2d(object[,] x)
        {
            if (!AreInputsValid(x))
            {
                return TestOutputs.Error.ToString();
            }

            string result = TestOutputs.Ok.ToString();
            for (int i = 0; i < x.GetLength(0); i++)
            {
                for (int j = 0; j < x.GetLength(1); j++)
                {
                    if (x[i, j].ToString().IgnoreCaseEquals(TestOutputs.Error))
                    {
                        return TestOutputs.Error.ToString();
                    }
                    
                    if (x[i, j].ToString().IgnoreCaseEquals(TestOutputs.Warning))
                    {
                        result = TestOutputs.Warning.ToString();
                    }
                    else
                    {
                        result = result.IgnoreCaseEquals(TestOutputs.Warning)
                            ? TestOutputs.Warning.ToString()
                            : TestOutputs.Ok.ToString();
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
        Name = "d.TestUtils_Not",
        Description = "Returns 'OK' if input is 'ERROR', 'ERROR' if 'OK' and 'WARNING' otherwise.",
        Category = "∂Excel: Test")]
    public static object Not(
        [ExcelArgument(Name = "Range", Description = "Input range.")]
        object[,] x)
    {
        object[,] output = new object[x.GetLength(0), x.GetLength(1)];
        
        if (!AreInputsValid(x))
        {
            for (int i = 0; i < output.GetLength(0); i++)
            {
                for (int j = 0; j < output.GetLength(1); j++)
                {
                    output[i, j] = TestOutputs.Error.ToString();
                } 
            }

            return output;
        }

        for (int i = 0; i < x.GetLength(0); i++)
        {
            for (int j = 0; j < x.GetLength(1); j++)
            {
                if (x[i, j].ToString().IgnoreCaseEquals(TestOutputs.Ok.ToString()))
                {
                    output[i, j] = TestOutputs.Error.ToString();
                }
                else if (x[i, j].ToString().IgnoreCaseEquals(TestOutputs.Error))
                {
                    output[i, j] = TestOutputs.Ok.ToString();
                }
                else
                {
                    output[i, j] = TestOutputs.Warning.ToString();
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
    /// <returns>'OK' if input range contains any 'OK's, otherwise 'WARNING' if there are no 'OK's but at least one
    /// 'WARNING', otherwise 'ERROR'.</returns>
    [ExcelFunction(
        Name = "d.TestUtils_Or",
        Description = "Acts like a fuzzy logic 'Or' with the following rules." +
        "\n  'OK' or X = 'OK'" +
        "\n  'WARNING' or 'ERROR' = 'WARNING'" +
        "\ni.e. 'OK' > 'WARNING' > 'ERROR' it can be seen as finding ANY 'OK's.",
        Category = "∂Excel: Test")]
    public static string Or(
        [ExcelArgument(Name = "Range", Description = "Input range.")]
        params object[] xRange)
    {
        object[] resultArray = new object[xRange.Length];

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
                resultArray[i] = xRange[i].ToString() ?? "";
            }
        }

        return Or1d(resultArray);

        string Or1d(object[] x)
        {
            if (!AreInputsValid(x))
            {
                return TestOutputs.Error.ToString();
            }

            string result = TestOutputs.Ok.ToString();
            for (int i = 0; i < x.GetLength(0); i++)
            {
                if (string.Compare(x[i].ToString(), TestOutputs.Ok.ToString(), StringComparison.OrdinalIgnoreCase) == 0)
                {
                    return TestOutputs.Ok.ToString();
                }
                
                if (string.Compare(
                        strA: x[i].ToString(), 
                        strB: TestOutputs.Warning.ToString(),
                        comparisonType: StringComparison.OrdinalIgnoreCase) == 0)
                {
                    result = TestOutputs.Warning.ToString();
                }
                else
                {
                    result = result == TestOutputs.Warning.ToString()
                        ? TestOutputs.Warning.ToString()
                        : TestOutputs.Error.ToString();
                }
            }
            
            return result;
        }

        string Or2d(object[,] x)
        {
            if (!AreInputsValid(x))
            {
                return TestOutputs.Error.ToString();
            }

            string result = TestOutputs.Ok.ToString();
            for (int i = 0; i < x.GetLength(0); i++)
            {
                for (int j = 0; j < x.GetLength(1); j++)
                {
                    if (string.Compare(
                            strA: x[i, j].ToString(), 
                            strB: TestOutputs.Ok.ToString(),
                            comparisonType: StringComparison.OrdinalIgnoreCase) == 0)
                    {
                        return TestOutputs.Ok.ToString();
                    }
                    
                    if (string.Compare(
                            strA: x[i, j].ToString(), 
                            strB: TestOutputs.Warning.ToString(),
                            comparisonType: StringComparison.OrdinalIgnoreCase) == 0)
                    {
                        result = TestOutputs.Warning.ToString();
                    }
                    else
                    {
                        result = result == TestOutputs.Warning.ToString()
                            ? TestOutputs.Warning.ToString()
                            : TestOutputs.Error.ToString();
                    }
                }
            }
            
            return result;
        }
    }
}
