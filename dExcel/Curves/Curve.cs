namespace dExcel.Curves;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using QLNet;

public static class Curve
{
    //[ExcelFunction(Name="d.Curve_Create")]
    //public static string Create(string handle, object[,] datesRange, object[,] discountFactorsRange)
    //{
    //    // TODO: Check datesRange and discountFactorsRange have same dimensions.
    //    List<Date> dates = new();
    //    List<double> discountFactors = new();
    //    for (int i = 0; i < datesRange.GetLength(0); i++)
    //    {
    //        dates.Add((Date)DateTime.FromOADate((double)datesRange[i, 0]));
    //        discountFactors.Add((double)discountFactorsRange[i, 0]);
    //    }

    //    //QLNet.PiecewiseYieldCurve<Discount, LogLinear> curve = new ()
    //    InterpolatedDiscountCurve<LogLinear> interpolatedDiscountCurve
    //        = new(dates, discountFactors, new Actual365Fixed(), new LogLinear());
    //    return DataObjectController.Add(handle, interpolatedDiscountCurve);
    //}

    //[ExcelFunction(Name = "d.Curve_DF")]
    //public static double GetDiscountFactor(string handle, DateTime date)
    //{
    //    return ((InterpolatedDiscountCurve<LogLinear>)DataObjectController.GetDataObject(handle)).discount(date);
    //}

    //[ExcelFunction(Name = "d.Curve_GetForwardRate")]

}

