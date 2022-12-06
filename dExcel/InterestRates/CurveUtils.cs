namespace dExcel.InterestRates;

using ExcelUtils;
using ExcelDna.Integration;
using QLNet;

public static class CurveUtils
{
    public static CurveDetails GetCurveDetails(string handle)
        => (CurveDetails)DataObjectController.GetDataObject(handle);

    /// <summary>
    /// Gets the curve object from a given handle which can be used to extract discount factors, zero rates etc.
    /// </summary>
    /// <param name="handle">The handle for the relevant curve object.</param>
    /// <returns>Returns the YieldTermStructure object.</returns>
    public static YieldTermStructure? GetCurveObject(string handle)
        => ((CurveDetails)DataObjectController.GetDataObject(handle)).TermStructure as YieldTermStructure;

    /// <summary>
    /// Gets the DayCounter object from a given handle which can be used to calculate year fractions.
    /// </summary>
    /// <param name="handle">The handle for the relevant curve object.</param>
    /// <returns>Returns the DayCounter object e.g. Actual365Fixed.</returns>
    private static DayCounter GetCurveDayCountConvention(string handle)
        => ((CurveDetails)DataObjectController.GetDataObject(handle)).DayCountConvention; 
    
    /// <summary>
    /// Gets the interpolation object from a given handle.
    /// </summary>
    /// <param name="handle">The handle for the relevant curve object.</param>
    /// <returns>Returns the interpolation object e.g. LogLinear.</returns>   
    private static IInterpolationFactory GetInterpolation(string handle)
        => ((CurveDetails)DataObjectController.GetDataObject(handle)).Interpolation;

    /// <summary>
    /// Creates a QLNet YieldTermStructure curve object which is stored in the DataObjectController.
    /// </summary>
    /// <param name="handle">Handle or name to extract curve from DataObjectController.</param>
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
            return CommonUtils.DExcelErrorMessage("Dates and discount factors have incompatible sizes: " +
                $"({datesRange.GetLength(0)} != {discountFactorsRange.GetLength(0)}).");
        }

        List<Date>? dates = new();
        List<double>? discountFactors = new();
        for (int i = 0; i < datesRange.GetLength(0); i++)
        {
            dates.Add((Date)DateTime.FromOADate((double)datesRange[i, 0]));
            discountFactors.Add((double)discountFactorsRange[i, 0]);
        }

        string? dayCountConventionParameter = ExcelTable.GetTableValue<string>(curveParameters, "Value", "DayCountConvention", 0);
        
        if (dayCountConventionParameter == null)
        {
            return CommonUtils.DExcelErrorMessage("'DayCountConvention' not set in parameters.");
        }

        DayCounter? dayCountConvention =
            dayCountConventionParameter.ToUpper() switch
            {
                "ACT360" or "ACTUAL360" => new Actual360(),
                "ACT365" or "ACTUAL365" => new Actual365Fixed(),
                "ACTACT" or "ACTUALACTUAL" => new ActualActual(),
                "BUSINESS252" => new Business252(),
                "30360" or "THIRTY360" => new Thirty360(Thirty360.Thirty360Convention.BondBasis, null),
                _ => null,
            };

        if (dayCountConvention == null)
        {
            return CommonUtils.DExcelErrorMessage($"Invalid 'DayCountConvention': {dayCountConventionParameter}");
        }

        string? interpolationParameter = ExcelTable.GetTableValue<string>(curveParameters, "Value", "Interpolation", 0);

        if (interpolationParameter == null)
        {
            return CommonUtils.DExcelErrorMessage("'Interpolation' not set in parameters.");
        }

        IInterpolationFactory? interpolation =
            interpolationParameter.ToUpper() switch
            {
                "BACKWARDFLAT" => new BackwardFlat(),
                "CUBIC" => new Cubic(),
                "FORWARDFLAT" => new ForwardFlat(),
                "LINEAR" => new Linear(),
                "LOGCUBIC" => new LogCubic(),
                "EXPONENTIAL" => new LogLinear(),
                _ => null,
            };

        if (interpolation == null)
        {
            return CommonUtils.DExcelErrorMessage($"Invalid 'interpolation' method: {interpolationParameter}");
        }

        string? calendarsParameter = ExcelTable.GetTableValue<string>(curveParameters, "Value", "Calendars");
        IEnumerable<string>? calendars = calendarsParameter?.Split(',').Select(x => x.ToString().Trim().ToUpper());
        Type interpolationType = typeof(InterpolatedDiscountCurve<>).MakeGenericType(interpolation.GetType());
        object? termStructure= Activator.CreateInstance(interpolationType, dates, discountFactors, dayCountConvention, interpolation);
        CurveDetails curveDetails = new(termStructure, dayCountConvention, interpolation, dates, discountFactors);
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
            Description = "The dates for which to get the discount factors.")]
            object[] dates)
    {
        object[,] discountFactors = new object[dates.Length, 1];
        YieldTermStructure? curve = GetCurveObject(handle);
        for (int i = 0; i < dates.Length; i++)
        {
            discountFactors[i, 0] = curve.discount((Date)DateTime.FromOADate((double)dates[i]));
        }
        
        return discountFactors;
    }

    [ExcelFunction(
        Name = "d.Curve_GetForwardRates",
        Description = "Gets forward rate from the curve for the given start and end date as well as the compounding convention.",
        Category = "∂Excel: Interest Rates")]
    public static object GetForwardRate(
        string handle, 
        object[,] startDatesRange, 
        object[,] endDatesRange, 
        string compoundingConvention)
    {
        YieldTermStructure? curve = GetCurveObject(handle);
        if (curve is null)
        {
            return CommonUtils.DExcelErrorMessage($"{handle} returned null object.");
        }
        
        (Compounding? compounding, Frequency? frequency)
            = compoundingConvention.ToUpper() switch
            {
                "SIMPLE" => (Compounding.Simple, Frequency.Once),
                "NACM" => (Compounding.Compounded, Frequency.Monthly),
                "NACQ" => (Compounding.Compounded, Frequency.Quarterly),
                "NACS" => (Compounding.Compounded, Frequency.Semiannual),
                "NACA" => (Compounding.Compounded, Frequency.Annual),
                "NACC" => (Compounding.Continuous, Frequency.NoFrequency),
                _ => ((Compounding?)null, (Frequency?)null),
            };

        if (compounding == null || frequency == null)
        {
            return CommonUtils.DExcelErrorMessage($"Invalid compounding convention: {compoundingConvention}");
        }

        List<DateTime> startDates = ArrayUtils.ConvertExcelRangeToList<DateTime>(startDatesRange, 0);
        List<DateTime> endDates = ArrayUtils.ConvertExcelRangeToList<DateTime>(endDatesRange, 0);
        DayCounter dayCountConvention = GetCurveDayCountConvention(handle);

        object[,] forwardRates = new object[startDates.Count, 1];
        for (int i = 0; i < startDates.Count; i++)
        {
            forwardRates[i, 0] = 
                curve.forwardRate((Date)startDates[i], (Date)endDates[i], dayCountConvention, (Compounding)compounding, (Frequency)frequency).rate();
        }

        return forwardRates;
    }

    /// <summary>
    /// Gets the zero rate(s) from a YieldTermStructure curve object for a given set of date(s).
    /// </summary>
    /// <param name="handle">The curve object handle (i.e., name).</param>
    /// <param name="datesRange">The range of dates.</param>
    /// <param name="compoundingConvention">The compounding convention.</param>
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
            Name = "(Optional)Compounding Convention",
            Description = "The compounding convention: Simple, NACC, NACM, NACQ, NACS, NACA \n" +
                          "Default = NACC")]
            string compoundingConvention = "NACC")
    {
        YieldTermStructure? curve = GetCurveObject(handle);
        List<Date> dates = new();
        DayCounter dayCountConvention = GetCurveDayCountConvention(handle);

        (Compounding? compounding, Frequency? frequency)
            = compoundingConvention.ToUpper() switch
            {
                "SIMPLE" => (Compounding.Simple, Frequency.Once),
                "NACM" => (Compounding.Compounded, Frequency.Monthly),
                "NACQ" => (Compounding.Compounded, Frequency.Quarterly),
                "NACS" => (Compounding.Compounded, Frequency.Semiannual),
                "NACA" => (Compounding.Compounded, Frequency.Annual),
                "NACC" => (Compounding.Continuous, Frequency.NoFrequency),
                _ => ((Compounding?)null, (Frequency?)null),
            };

        object[,] zeroRates = new object[datesRange.Length, 1];
        if (compounding == null)
        {
            zeroRates[0, 0] = CommonUtils.DExcelErrorMessage($"Invalid compounding convention '{compoundingConvention}'.");
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
