namespace dExcel;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

public static class FuzzyLogic
{
    public enum FuzzyOutputs
    {
        OK, ERROR, WARNING,
    }

    [ExcelFunction(
        Name = "d.Equal",
        Description = "",
        Category = "∂Excel: Mathematics")]
    public static string Equal(object a, object b)
    {
        if (double.TryParse(a.ToString(), out var x) && double.TryParse(b.ToString(), out var y))
        {
            return Math.Abs(y - x) < 0.00000000001 ? FuzzyOutputs.OK.ToString() : FuzzyOutputs.ERROR.ToString();   
        }
        
        return string.Compare(a.ToString(), b.ToString(), true) == 0? FuzzyOutputs.OK.ToString() : FuzzyOutputs.ERROR.ToString();
    }

    [ExcelFunction(
        Name = "d.GreaterThan",
        Description = "",
        Category = "∂Excel: Mathematics")]
    public static string GreaterThan(double a, double b) =>
        (double)a > (double)b ? FuzzyOutputs.OK.ToString() : FuzzyOutputs.ERROR.ToString();

    [ExcelFunction(
        Name = "d.LessThan",
        Description = "",
        Category = "∂Excel: Mathematics")]
    public static string LessThan(double a, double b) =>
        (double)a > (double)b ? FuzzyOutputs.OK.ToString() : FuzzyOutputs.ERROR.ToString();

    [ExcelFunction(
        Name = "d.Or",
        Description = "",
        Category = "∂Excel: Mathematics")]
    public static string Or(params object[] x)
    {
        if (x.Contains(FuzzyOutputs.OK.ToString()))
        {
            return FuzzyOutputs.OK.ToString();
        }
        else if (x.Contains(FuzzyOutputs.WARNING.ToString()))
        {
            return FuzzyOutputs.WARNING.ToString();
        }
        else
        {
            return FuzzyOutputs.ERROR.ToString();
        }
    }

    [ExcelFunction(
    Name = "d.And",
    Description = "",
    Category = "∂Excel: Mathematics")]
    public static string And(params object[] x)
    {
        if (x.Contains(FuzzyOutputs.ERROR.ToString()))
        {
            return FuzzyOutputs.ERROR.ToString();
        }
        else if (x.Contains(FuzzyOutputs.WARNING.ToString()))
        {
            return FuzzyOutputs.WARNING.ToString();
        }
        else
        {
            return FuzzyOutputs.OK.ToString();
        }
    }


    [ExcelFunction(
    Name = "d.Not",
    Description = "",
    Category = "∂Excel: Mathematics")]
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
}
