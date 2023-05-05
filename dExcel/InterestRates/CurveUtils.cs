﻿using dExcel.Dates;
using dExcel.ExcelUtils;
using dExcel.Utilities;
using ExcelDna.Integration;
using QL = QuantLib;

namespace dExcel.InterestRates;

/// <summary>
/// A collection of utility functions for dealing with interest rate curves.
/// </summary>
public static class CurveUtils
{
    public static CurveDetails GetCurveDetails(string handle)
    {
        DataObjectController controller = DataObjectController.Instance;
        return (CurveDetails)controller.GetDataObject(handle);
    }

    /// <summary>
    /// Gets the curve object from a given handle which can be used to extract discount factors, zero rates etc.
    /// </summary>
    /// <param name="handle">The handle for the relevant curve object.</param>
    /// <returns>Returns the YieldTermStructure object.</returns>
    public static QL.YieldTermStructure? GetCurveObject(string handle)
    {
        DataObjectController controller = DataObjectController.Instance;
        QL.YieldTermStructure? curve = ((CurveDetails)controller.GetDataObject(handle)).TermStructure as QL.YieldTermStructure;
        QL.Settings.instance().setEvaluationDate(curve?.referenceDate());
        return curve;
    }

    /// <summary>
    /// Gets the DayCounter object from a given handle which can be used to calculate year fractions.
    /// </summary>
    /// <param name="handle">The handle for the relevant curve object.</param>
    /// <returns>Returns the DayCounter object e.g. Actual365Fixed.</returns>
    private static QL.DayCounter GetCurveDayCountConvention(string handle)
    {
        DataObjectController dataObjectController = DataObjectController.Instance;
        return ((CurveDetails)dataObjectController.GetDataObject(handle)).DayCountConvention; 
    }

    /// <summary>
    /// Gets the interpolation object from a given handle.
    /// </summary>
    /// <param name="handle">The handle for the relevant curve object.</param>
    /// <returns>Returns the interpolation object e.g. LogLinear.</returns>   
    private static string GetInterpolation(string handle)
    {
        DataObjectController dataObjectController = DataObjectController.Instance;
        return ((CurveDetails)dataObjectController.GetDataObject(handle)).DiscountFactorInterpolation;
    }

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
        Category = "∂Excel: Interest Rates",
        IsVolatile = true)]
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
                $"({datesRange.GetLength(0)} ≠ {discountFactorsRange.GetLength(0)}).");
        }

        List<QL.Date> dates = new();
        List<double> discountFactors = new();
        for (int i = 0; i < datesRange.GetLength(0); i++)
        {
            dates.Add(DateTime.FromOADate((double)datesRange[i, 0]).ToQuantLibDate());
            discountFactors.Add((double)discountFactorsRange[i, 0]);
        }

        string? dayCountConventionParameter = ExcelTableUtils.GetTableValue<string>(curveParameters, "Value", "DayCountConvention", 0);
        if (dayCountConventionParameter == null)
        {
            return CommonUtils.DExcelErrorMessage("Parameter not set: 'DayCountConvention'");
        }

        if (!ParserUtils.TryParseQuantLibDayCountConvention(
                dayCountConventionToParse: dayCountConventionParameter, 
                dayCountConvention: out QL.DayCounter? dayCountConvention,
                errorMessage: out string? dayCountConventionErrorMessage))
        {
            return dayCountConventionErrorMessage;
        }

        string? interpolationParameter = ExcelTableUtils.GetTableValue<string>(curveParameters, "Value", "Interpolation", 0);
        if (interpolationParameter == null)
        {
            return CommonUtils.DExcelErrorMessage("'Interpolation' not set in parameters.");
        }

        string? calendarsParameter = ExcelTableUtils.GetTableValue<string>(curveParameters, "Value", "Calendars", 0);
        IEnumerable<string>? calendars = calendarsParameter?.Split(',').Select(x => x.ToString().Trim().ToUpper());
        (QL.Calendar? calendar, string errorMessage) = DateUtils.ParseCalendars(calendarsParameter);  

        if (string.Compare(interpolationParameter, "EXPONENTIAL", StringComparison.OrdinalIgnoreCase) == 0)
        {
            QL.DiscountCurve discountCurve = 
                new(
                    dates: new QL.DateVector(dates), 
                    discounts: new QL.DoubleVector(discountFactors), 
                    dayCounter: dayCountConvention, 
                    calendar: calendar);

            CurveDetails curveDetails = 
                new(
                    termStructure: discountCurve, 
                    dayCountConvention: dayCountConvention, 
                    interpolation: interpolationParameter, 
                    discountFactorDates: dates.Select(x => x.ToDateTime()), 
                    discountFactors: discountFactors);
            
            DataObjectController dataObjectController = DataObjectController.Instance;
            return dataObjectController.Add(handle, curveDetails);
        }

        if (
            string.Compare(interpolationParameter, "CUBIC", StringComparison.OrdinalIgnoreCase) == 0)
        {
            QL.NaturalCubicDiscountCurve discountCurve = 
                new(
                    dates: new QL.DateVector(dates), 
                    discounts: new QL.DoubleVector(discountFactors), 
                    dayCounter: dayCountConvention, 
                    calendar: calendar);
            
            CurveDetails curveDetails = 
                new(
                    termStructure: discountCurve, 
                    dayCountConvention: dayCountConvention, 
                    interpolation: interpolationParameter, 
                    discountFactorDates: dates.Select(x => x.ToDateTime()), 
                    discountFactors: discountFactors);
            
            DataObjectController dataObjectController = DataObjectController.Instance;
            return dataObjectController.Add(handle, curveDetails);
        }

        
        if (interpolationParameter.ToLower() == "LOGCUBIC")
        {
            QL.MonotonicLogCubicDiscountCurve discountCurve = 
                new(
                    dates: new QL.DateVector(dates), 
                    discounts: new QL.DoubleVector(discountFactors), 
                    dayCounter: dayCountConvention, 
                    calendar: calendar);
            
            CurveDetails curveDetails = 
                new(
                    termStructure: discountCurve, 
                    dayCountConvention: dayCountConvention, 
                    interpolation: interpolationParameter, 
                    discountFactorDates: dates.Select(x => x.ToDateTime()), 
                    discountFactors: discountFactors);
            
            DataObjectController dataObjectController = DataObjectController.Instance;
            return dataObjectController.Add(handle, curveDetails);
        }
        
        return CommonUtils.DExcelErrorMessage($"Unsupported interpolation method: {interpolationParameter}");
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
        Category = "∂Excel: Interest Rates",
        IsVolatile = true)]
    public static object GetDiscountFactors(
        [ExcelArgument(
            Name = "Handle",
            Description = "The 'handle' or name used to store and retrieve the curve.")]
            string handle,
        [ExcelArgument(
            Name = "Dates/Year Fractions",
            Description = "The dates/year fractions for which to get the discount factors.")]
            object[] dates)
    {
        object[,] discountFactors = new object[dates.Length, 1];
        QL.YieldTermStructure? curve = GetCurveObject(handle);
        // Most "current" dates are in the range 40,000 - 50,000.
        // Hence it's safe to assume that if it's less than 1,000 it must be a year fraction.
        if ((double) dates[0] < 1_000)
        {
            for (int i = 0; i < dates.Length; i++)
            {
                discountFactors[i, 0] = curve.discount((double) dates[i]);
            }
        }
        else
        {
            for (int i = 0; i < dates.Length; i++)
            {
                discountFactors[i, 0] = curve.discount(DateTime.FromOADate((double)dates[i]).ToQuantLibDate());
            }    
        }
        
        return discountFactors;
    }

    [ExcelFunction(
        Name = "d.Curve_GetForwardRates",
        Description = "Gets forward rate from the curve for the given start and end date as well as the compounding convention.",
        Category = "∂Excel: Interest Rates")]
    public static object GetForwardRates(
        string handle, 
        object[,] startDatesRange, 
        object[,] endDatesRange, 
        string compoundingConventionParameter)
    {
        QL.YieldTermStructure? curve = GetCurveObject(handle);
        if (curve is null)
        {
            return CommonUtils.DExcelErrorMessage($"{handle} returned null object.");
        }

        if (!ParserUtils.TryParseQuantLibCompoundingConvention(
                compoundingConventionParameter,
                out (QL.Compounding compounding, QL.Frequency frequency)? compoundingConvention,
                out string? compoundingConventionErrorMessage))
        {
            return compoundingConventionErrorMessage;
        }
        

        List<DateTime> startDates = ExcelArrayUtils.ConvertExcelRangeToList<DateTime>(startDatesRange);
        List<DateTime> endDates = ExcelArrayUtils.ConvertExcelRangeToList<DateTime>(endDatesRange);
        QL.DayCounter dayCountConvention = GetCurveDayCountConvention(handle);

        object[,] forwardRates = new object[startDates.Count, 1];
        for (int i = 0; i < startDates.Count; i++)
        {
            forwardRates[i, 0] = 
                curve.forwardRate(
                    startDates[i].ToQuantLibDate(), 
                    endDates[i].ToQuantLibDate(),
                    dayCountConvention, 
                    compoundingConvention.Value.compounding, 
                    compoundingConvention.Value.frequency).rate();
        }

        return forwardRates;
    }

    /// <summary>
    /// Gets the zero rate(s) from a YieldTermStructure curve object for a given set of date(s).
    /// </summary>
    /// <param name="handle">The curve object handle (i.e., name).</param>
    /// <param name="datesRange">The range of dates.</param>
    /// <param name="compoundingConventionParameter">The compounding convention.</param>
    /// <returns>The zero rate(s) for the given date(s).</returns>
    [ExcelFunction(
        Name = "d.Curve_GetZeroRates",
        Description = "Gets the zero rate(s) from a curve object for a given set of date(s).",
        Category = "∂Excel: Interest Rates",
        HelpTopic = "https://wiki.fsa-aks.deloitte.co.za/doku.php?id=valuations:methodology:curves_and_bootstrapping:interest_rate_calculations")]
    public static object GetZeroRates(
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
            Description = "The compounding convention: Simple, NACC, NACM, NACQ, NACS, NACA \nDefault = NACC")]
            string compoundingConventionParameter = "NACC")
    {
        QL.YieldTermStructure? curve = GetCurveObject(handle);
        if (curve is null)
        {
            return CommonUtils.DExcelErrorMessage($"Curve with handle {handle} not found. Try refreshing it.");
        }

        List<QL.Date> dates = new();
        QL.DayCounter dayCountConvention = GetCurveDayCountConvention(handle);

        if (!ParserUtils.TryParseQuantLibCompoundingConvention(
                compoundingConventionParameter,
                out (QL.Compounding compounding, QL.Frequency frequency)? compoundingConvention,
                out string? compoundingConventionErrorMessage))
        {
            return compoundingConventionErrorMessage; 
        }
        
        object[,] zeroRates = new object[datesRange.Length, 1];
        for (int i = 0; i < datesRange.GetLength(0); i++)
        {
            dates.Add(DateTime.FromOADate((double)datesRange[i, 0]).ToQuantLibDate());
            zeroRates[i, 0] = 
                curve.zeroRate(
                    dates[i], 
                    dayCountConvention, 
                    compoundingConvention.Value.compounding, 
                    compoundingConvention.Value.frequency).rate();
        }

        return zeroRates;
    }

    /// <summary>
    /// Extracts the instruments used to bootstrap the curve.
    /// </summary>
    /// <param name="handle">The 'handle' or name used to refer to the object in memory. Each object must have a unique
    /// handle.</param>
    /// <returns>A 2d object containing the list of instruments used to bootstrap the curve. If the curve wasn't
    /// bootstrapped it just returns a warning message.</returns>
    [ExcelFunction(
        Name = "d.Curve_GetInstruments",
        Description = "Extracts the instruments used to bootstrap the curve.",
        Category = "∂Excel: Interest Rates")]
    public static object GetInstruments(
        [ExcelArgument(
            Name = "Handle", 
            Description = 
                "The 'handle' or name used to refer to the object in memory.\n" + 
                "Each object must have a unique handle.")]
        string handle)
    {
        CurveDetails curve = GetCurveDetails(handle);
        List<object> instrumentGroups = curve.InstrumentGroups.ToList();

        if (instrumentGroups.Count == 0)
        {
            return CommonUtils.DExcelErrorMessage(
                "No instruments found. Was this bootstrapped or built from discount factors directly?");
        }
        
        int numberOfRows = 0;
        int maxColumnNumber = 0;
        foreach (object[,] instrumentGroup in instrumentGroups)
        {
            numberOfRows += instrumentGroup.GetLength(0) + 1;
            maxColumnNumber = Math.Max(maxColumnNumber, instrumentGroup.GetLength(1));
        }
        
        object[,] output = new object[numberOfRows - 1, maxColumnNumber];

        int row = 0;
        foreach (object[,] instrumentGroup in instrumentGroups)
        {
            for (int i = 0; i < instrumentGroup.GetLength(0); i++)
            {
                for (int j = 0; j < instrumentGroup.GetLength(1); j++)
                {
                    if (instrumentGroup[i, j] == null)
                    {
                        output[row, j] = "";
                    }
                    else if (instrumentGroup[i, j].ToString() == ExcelEmpty.Value.ToString())
                    {
                        output[row, j] = "";
                    }
                    else
                    {
                        output[row, j] = instrumentGroup[i, j];
                    }
                }

                for (int j = instrumentGroup.GetLength(1); j < maxColumnNumber; j++)
                {
                    output[row, j] = "";
                }
                
                row++;
            }

            if (row != numberOfRows - 1)
            {
                for (int j = 0; j < maxColumnNumber; j++)
                {
                    output[row, j] = "";
                }
            }
            
            row++;
        }
        
        return output; 
    }
}
