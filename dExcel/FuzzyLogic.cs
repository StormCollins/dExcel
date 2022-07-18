namespace dExcel;

using System;
using System.Linq;
using ExcelDna.Integration;

public static class FuzzyLogicUtils
{
    public enum FuzzyOutputs
    {
        OK, ERROR, WARNING,
    }

    [ExcelFunction(
        Name = "d.Logic_Equal",
        Description = "Returns 'Okay' if two values are equal, otherwise it returns 'Error'.",
        Category = "∂Excel: Logic")]
    public static string Equal(object a, object b)
    {
        if (double.TryParse(a.ToString(), out var x) && double.TryParse(b.ToString(), out var y))
        {
            return Math.Abs(y - x) < 0.00000000001 ? FuzzyOutputs.OK.ToString() : FuzzyOutputs.ERROR.ToString();   
        }
        
        return string.Compare(a.ToString(), b.ToString(), true) == 0? FuzzyOutputs.OK.ToString() : FuzzyOutputs.ERROR.ToString();
    }

    [ExcelFunction(
        Name = "d.Logic_GreaterThan",
        Description = "Returns 'Okay' if one value is greater than the other, otherwise it returns 'Error'.",
        Category = "∂Excel: Mathematics")]
    public static string GreaterThan(double a, double b) =>
        (double)a > (double)b ? FuzzyOutputs.OK.ToString() : FuzzyOutputs.ERROR.ToString();

    [ExcelFunction(
        Name = "d.Logic_LessThan",
        Description = "Returns 'Okay' if one value is less than the other, otherwise it returns 'Error'.",
        Category = "∂Excel: Logic")]
    public static string LessThan(double a, double b) =>
        (double)a > (double)b ? FuzzyOutputs.OK.ToString() : FuzzyOutputs.ERROR.ToString();

    [ExcelFunction(
        Name = "d.Logic_Or",
        Description = "",
        Category = "∂Excel: Logic")]
    public static string Or(params object[] xRange)
    {
        var resultArray = new string[xRange.Length];

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
                resultArray[i] = xRange[i].ToString();
            }
        }

        return Or1d(resultArray);

        string Or1d(object[] x)
        {
            var result = FuzzyOutputs.OK.ToString();
            for (int i = 0; i < x.GetLength(0); i++)
            {
                if (string.Compare(x[i].ToString(), FuzzyOutputs.OK.ToString(), true) == 0)
                {
                    return FuzzyOutputs.OK.ToString();
                }
                else if (string.Compare(x[i].ToString(), FuzzyOutputs.WARNING.ToString(), true) == 0)
                {
                    result = FuzzyOutputs.WARNING.ToString();
                }
                else
                {
                    result = (result == FuzzyOutputs.WARNING.ToString()) ? FuzzyOutputs.WARNING.ToString() : FuzzyOutputs.ERROR.ToString();
                }
            }

            return result;
        }

        string Or2d(object[,] x)
        {
            var result = FuzzyOutputs.OK.ToString();
            for (int i = 0; i < x.GetLength(0); i++)
            {
                for (int j = 0; j < x.GetLength(1); j++)
                {
                    if (string.Compare(x[i, j].ToString(), FuzzyOutputs.OK.ToString(), true) == 0)
                    {
                        return FuzzyOutputs.OK.ToString();
                    }
                    else if (string.Compare(x[i, j].ToString(), FuzzyOutputs.WARNING.ToString(), true) == 0)
                    {
                        result = FuzzyOutputs.WARNING.ToString();
                    }
                    else
                    {
                        result = (result == FuzzyOutputs.WARNING.ToString()) ? FuzzyOutputs.WARNING.ToString() : FuzzyOutputs.ERROR.ToString();
                    }
                }
            }

            return result;
        }
    }

    [ExcelFunction(
        Name = "d.Logic_And",
        Description = "",
        Category = "∂Excel: Logic")]
    public static string And(params object[] xRange)
    {
        var resultArray = new string[xRange.Length];

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
                resultArray[i] = xRange[i].ToString();
            }
        }

        return And1d(resultArray);
    
        string And1d(object[] x)
        {
            var result = FuzzyOutputs.OK.ToString();
            for (int i = 0; i < x.GetLength(0); i++)
            {
                if (string.Compare(x[i].ToString(), FuzzyOutputs.ERROR.ToString(), true) == 0)
                {
                    return FuzzyOutputs.ERROR.ToString();
                }
                else if (string.Compare(x[i].ToString(), FuzzyOutputs.WARNING.ToString(), true) == 0)
                {
                    result = FuzzyOutputs.WARNING.ToString();
                }
                else
                {
                    result = (result == FuzzyOutputs.WARNING.ToString()) ? FuzzyOutputs.WARNING.ToString() : FuzzyOutputs.OK.ToString();
                }
            }

            return result;
        }

        string And2d(object[,] x)
        {
            var result = FuzzyOutputs.OK.ToString();
            for (int i = 0; i < x.GetLength(0); i++)
            {
                for (int j = 0; j < x.GetLength(1); j++)
                {
                    if (string.Compare(x[i, j].ToString(), FuzzyOutputs.ERROR.ToString(), true) == 0)
                    {
                        return FuzzyOutputs.ERROR.ToString();
                    }
                    else if (string.Compare(x[i, j].ToString(), FuzzyOutputs.WARNING.ToString(), true) == 0)
                    {
                        result = FuzzyOutputs.WARNING.ToString();
                    }
                    else
                    {
                        result = (result == FuzzyOutputs.WARNING.ToString()) ? FuzzyOutputs.WARNING.ToString() : FuzzyOutputs.OK.ToString();
                    }
                }
            }

            return result;
        }
    }

    [ExcelFunction(
    Name = "d.Logic_Not",
    Description = "",
    Category = "∂Excel: Logic")]
    public static object Not(object[] x)
    {
        var output = new object[x.Length];
        for (int i = 0; i < x.Length; i++)
        {
            output[i] =
                x[i].ToString() == FuzzyOutputs.OK.ToString() ?
                    FuzzyOutputs.ERROR.ToString() : 
                    x[i].ToString() == FuzzyOutputs.ERROR.ToString() ?
                        FuzzyOutputs.OK.ToString() : FuzzyOutputs.WARNING.ToString();
        }
        return output;
    }

    [ExcelFunction(
        Name = "d.Logic_IsTrue",
        Description = "Returns 'Okay' if parameter is true, otherwise returns 'Error'.",
        Category = "∂Excel: Logic")]
    public static object IsTrue(object x)
    {
        if ((bool)x)
        {
            return FuzzyOutputs.OK.ToString();
        }
        else
        {
            return FuzzyOutputs.ERROR.ToString();
        }
    }

    [ExcelFunction(
        Name = "d.Logic_IsFalse",
        Description = "Returns 'Okay' if parameter is false, otherwise returns 'Error'.",
        Category = "∂Excel: Logic")]
    public static object IsFalse(object x)
    {
        if (!(bool)x)
        {
            return FuzzyOutputs.OK.ToString();
        }
        else
        {
            return FuzzyOutputs.ERROR.ToString();
        }
    }
}
