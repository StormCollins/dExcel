namespace dExcel.Curves;

using System;
using System.Collections.Generic;
using ExcelDna.Integration;
using ExcelDna.Registration;
using ExcelUtils;
using QLNet;

public static class Curve
{
    /// <summary>
    /// Gets the curve object from a given handle which can be used to extract discount factors, zero rates etc.
    /// </summary>
    /// <param name="handle">The handle for the relevant curve object.</param>
    /// <returns>Returns the YieldTermStructure object.</returns>
    private static YieldTermStructure GetCurveObject(string handle)
        => (YieldTermStructure)((Dictionary<string, object>)DataObjectController.GetDataObject(handle))["Curve.Object"];

    /// <summary>
    /// Gets the DayCounter object from a given handle which can be used to calculate year fractions.
    /// </summary>
    /// <param name="handle">The handle for the relevant curve object.</param>
    /// <returns>Returns the DayCounter object e.g. Actual365Fixed.</returns>
    private static DayCounter GetCurveDayCountConvention(string handle)
        => (DayCounter)((Dictionary<string, object>)DataObjectController.GetDataObject(handle))["Curve.DayCountConvention"]; 
    
    /// <summary>
    /// Gets the interpolation object from a given handle.
    /// </summary>
    /// <param name="handle">The handle for the relevant curve object.</param>
    /// <returns>Returns the interplation object e.g. LogLinear.</returns>   
    private static IInterpolationFactory GetInteroplation(string handle)
        => (IInterpolationFactory)((Dictionary<string, object>)DataObjectController.GetDataObject(handle))["Curve.Interpolation"];

    /// <summary>
    /// Creates a QLNet YieldTermStructure curve object which is stored in the DataObjectController.
    /// </summary>
    /// <param name="handle">Handle or name to extract curve from DataObjectContronller.</param>
    /// <param name="curveParameters">The parameters for curve construction e.g. interpolation, day count convention etc.</param>
    /// <param name="datesRange">The dates for the corresponding discount factors.</param>
    /// <param name="discountFactorsRange">The discount factors for the corresponding dates.</param>
    /// <returns>A string containing the handle and time stamp.</returns>
    [ExcelFunction(
        Name = "d.Curve_Create",
        Description = "Creates an interest rate curve given dates and corresponding discount factors.",
        Category = "∂Excel: Interest Rates")]
    public static string Create(
        [ExcelArgument(
            Name = "Handle",
            Description = "The 'handle' or name used to store & retrieve the curve.")]
            string handle,
        [ExcelArgument(
            Name = "Parameters",
            Description = "The parameters for curve construction e.g. interpolation, day count convention etc.")]
            object[,] curveParameters,
        [ExcelArgument(
            Name = "Dates", 
            Description = "The dates for the corresponding discount factors.")]
            object[,] datesRange,
        [ExcelArgument(
            Name = "Discount Factors",
            Description = "The discount factors for the corresponding dates.")]
            object[,] discountFactorsRange)
    {
        if (datesRange.GetLength(0) != discountFactorsRange.GetLength(0))
        {
            return $"#Error: Dates and discount factors have incompatible sizes " +
                $"({datesRange.GetLength(0)} & {discountFactorsRange.GetLength(0)}).";
        }

        List<Date> dates = new();
        List<double> discountFactors = new();
        for (int i = 0; i < datesRange.GetLength(0); i++)
        {
            dates.Add((Date)DateTime.FromOADate((double)datesRange[i, 0]));
            discountFactors.Add((double)discountFactorsRange[i, 0]);
        }

        string? dayCountConventionParameter = ExcelTable.LookupTableValue<string>(curveParameters, "Value", "DayCountConvention");
        
        if (dayCountConventionParameter == null)
        {
            return "DayCountConvention not set in parameters.";
        }

        DayCounter? dayCountConvention =
            dayCountConventionParameter.ToUpper() switch
            {
                "ACT360" or "ACTUAL360" => new Actual360(),
                "ACT365" or "ACTUAL365" => new Actual365Fixed(),
                "ACTACT" or "ACTUALACTUAL" => new ActualActual(),
                "BUSINESS252" => new Business252(),
                "30360" or "THIRTY360" => new Thirty360(),
                _ => null,
            };

        if (dayCountConvention == null)
        {
            return $"DayCountConvention '{dayCountConventionParameter}' invalid.";
        }

        string? interpolationParameter = ExcelTable.LookupTableValue<string>(curveParameters, "Value", "Interpolation");

        if (interpolationParameter == null)
        {
            return "Interpolation not set in parameters.";
        }

        IInterpolationFactory? interpolation =
            interpolationParameter.ToUpper() switch
            {
                "BACKWARDFLAT" => new BackwardFlat(),
                "CUBIC" => new Cubic(),
                "FORWARDFLAT" => new ForwardFlat(),
                "LINEAR" => new Linear(),
                "LOGCUBIC" => new LogCubic(),
                "LOGLINEAR" => new LogLinear(),
                _ => null,
            };

        if (interpolation == null)
        {
            return $"Interpolation '{interpolationParameter}' invalid.";
        }

        string? calendarsParameter = ExcelTable.LookupTableValue<string>(curveParameters, "Value", "Calendars");
        var calendars = calendarsParameter?.Split(',').Select(x => x.ToString().Trim().ToUpper());

        var interpolationType = typeof(InterpolatedDiscountCurve<>).MakeGenericType(interpolation.GetType());
        var termStructure= Activator.CreateInstance(interpolationType, dates, discountFactors, dayCountConvention, interpolation);
        Dictionary<string, object> curveDetails = new()
        {
            ["Curve.Object"] = termStructure,
            ["Curve.DayCountConvention"] = dayCountConvention,
            ["Curve.Interpolation"] = interpolation,
        };
        
        return DataObjectController.Add(handle, curveDetails);
    }

    public static InterpolatedDiscountCurve<LogLinear> GetDiscountCurve(string handle)
    {
        return (InterpolatedDiscountCurve<LogLinear>)DataObjectController.GetDataObject(handle);
    }

    /// <summary>
    /// Gets the discount factor(s) from a QLNet YieldTermStructure curve object for a given set of date(s).
    /// </summary>
    /// <param name="handle"></param>
    /// <param name="dates"></param>
    /// <returns></returns>
    [ExcelFunction(
        Name = "d.Curve_GetDiscountFactors",
        Description = "Gets the discount factor(s) from a curve object for a given set of date(s).",
        Category = "∂Excel: Interest Rates")]
    public static object[,] GetDiscountFactors(
        [ExcelArgument(
            Name = "Handle",
            Description = "The 'handle' or name used to store & retrieve the curve.")]
            string handle,
        [ExcelArgument(
            Name = "Dates",
            Description = "The dates for which to get the disocunt factors.")]
            object[] dates)
    {
        var discountFactors = new object[dates.Length, 1];
        var curve = GetCurveObject(handle);
        for (int i = 0; i < dates.Length; i++)
        {
            discountFactors[i, 0] = curve.discount((Date)DateTime.FromOADate((double)dates[i]));
        }
        
        return discountFactors;
    }

    [ExcelFunction(
        Name = "d.Curve_GetForwardRates",
        Description = "",
        Category = "∂Excel: Interest Rates")]
    public static double GetForwardRate(string handle, DateTime startDate, DateTime endDate)
    {
        var curve = GetCurveObject(handle);
        return curve.forwardRate((Date)startDate, (Date)endDate, new Actual365Fixed(), Compounding.Continuous).rate();
    }

    /// <summary>
    /// Gets the zero rate(s) from a YieldTermStructure curve object for a given set of date(s).
    /// </summary>
    /// <param name="handle">The curve object handle (i.e. name).</param>
    /// <param name="datesRange">The range of dates.</param>
    /// <returns>The zero rate(s) for the given date(s).</returns>
    [ExcelFunction(
        Name = "d.Curve_GetZeroRates",
        Description = "Gets the zero rate(s) from a curve object for a given set of date(s).",
        Category = "∂Excel: Interest Rates",
        HelpTopic = "https://wiki.fsa-aks.deloitte.co.za/doku.php?id=valuations:methodology:curves_and_bootstrapping:interest_rate_calculations")]
    public static object[,] GetZeroRates(
        [ExcelArgument(
            Name = "Handle",
            Description = "The 'handle' or name used to store & retrieve the curve.")]
            string handle,
        [ExcelArgument(
            Name = "Dates",
            Description = "The dates for which to calculate the zero rates.")]
            object[,] datesRange,
        [ExcelArgument(
            Name = "(Optional)Compounding",
            Description = "The compounding convention: NACC, NACM, NACQ, NACS, NACA \nDefault = NACC")]
            string compoundingConvention = "NACC")
    {
        var curve = GetCurveObject(handle);
        List<Date> dates = new();
        var dayCountConvention = GetCurveDayCountConvention(handle);

        (Compounding? compounding, Frequency? frequency)
            = compoundingConvention.ToUpper() switch
            {
                "NACM" => (Compounding.Compounded, Frequency.Monthly),
                "NACQ" => (Compounding.Compounded, Frequency.Quarterly),
                "NACS" => (Compounding.Compounded, Frequency.Semiannual),
                "NACA" => (Compounding.Compounded, Frequency.Annual),
                "NACC" => (Compounding.Continuous, Frequency.NoFrequency),
                _ => ((Compounding?)null, (Frequency?)null),
            };

        var zeroRates = new object[datesRange.Length, 1];
        if (compounding == null)
        {
            zeroRates[0, 0] = $"#ERROR: Invalid compounding convention '{compoundingConvention}'.";
            return zeroRates;
        }

        for (int i = 0; i < datesRange.GetLength(0); i++)
        {
            dates.Add((Date)DateTime.FromOADate((double)datesRange[i, 0]));
            zeroRates[i, 0] = curve.zeroRate(dates[i], dayCountConvention, compounding ?? Compounding.Continuous, frequency ?? Frequency.NoFrequency).rate();
        }

        return zeroRates;
    }
}
