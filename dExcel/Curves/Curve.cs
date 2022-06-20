namespace dExcel.Curves;

using System;
using System.Collections.Generic;
using ExcelDna.Integration;
using QLNet;

public static class Curve
{
    [ExcelFunction(
        Name = "d.Curve_Create",
        Description = "",
        Category = "∂Excel: Interest Rates")]
    public static string Create(string handle, object[,] datesRange, object[,] discountFactorsRange)
    {
        // TODO: Check datesRange and discountFactorsRange have same dimensions.
        List<Date> dates = new();
        List<double> discountFactors = new();
        for (int i = 0; i < datesRange.GetLength(0); i++)
        {
            dates.Add((Date)DateTime.FromOADate((double)datesRange[i, 0]));
            discountFactors.Add((double)discountFactorsRange[i, 0]);
        }

        //QLNet.PiecewiseYieldCurve<Discount, LogLinear> curve = new ()
        InterpolatedDiscountCurve<LogLinear> termStructure
            = new(dates, discountFactors, new Actual365Fixed(), new LogLinear());
        
        Dictionary<string, object> curveDetails = new()
        {
            ["Curve"] = termStructure,
        };
        
        return DataObjectController.Add(handle, curveDetails);
    }

    public static InterpolatedDiscountCurve<LogLinear> GetDiscountCurve(string handle)
    {
        return (InterpolatedDiscountCurve<LogLinear>)DataObjectController.GetDataObject(handle);
    }
    
    
    [ExcelFunction(
        Name = "d.Curve_DF",
        Description = "",
        Category = "∂Excel: Interest Rates")]
    public static object[,] GetDiscountFactors(string handle, object[] dates)
    {
        var discountFactors = new object[dates.Length, 1];

        var curve =
            (YieldTermStructure)((Dictionary<string, object>)DataObjectController.GetDataObject(handle))["Curve"];
        for (int i = 0; i < dates.Length; i++)
        {
            discountFactors[i, 0] = curve.discount((Date)DateTime.FromOADate((double)dates[i]));
        }
        
        return discountFactors;
    }

    [ExcelFunction(
        Name = "d.Curve_GetForwardRate",
        Description = "",
        Category = "∂Excel: Interest Rates")]
    public static double GetForwardRate(string handle, DateTime startDate, DateTime endDate)
    {
        var curve = ((InterpolatedDiscountCurve<LogLinear>)DataObjectController.GetDataObject(handle));
        return curve.forwardRate((Date)startDate, (Date)endDate, new Actual365Fixed(), Compounding.Continuous).rate();
    }

    [ExcelFunction(
        Name = "d.Curve_GetZeroRate",
        Description = "",
        Category = "∂Excel: Interest Rates")]
    public static double GetZeroRate(string handle, DateTime startDate, DateTime endDate)
    {
        var curve = ((InterpolatedDiscountCurve<LogLinear>)DataObjectController.GetDataObject(handle));
        return curve.zeroRate((Date)startDate, new Actual365Fixed(), Compounding.Continuous).rate();
    }
}
