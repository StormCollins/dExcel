using dExcel.CommonEnums;
using dExcel.Dates;
using dExcel.ExcelUtils;
using dExcel.Utilities;
using ExcelDna.Integration;
using QL = QuantLib;

namespace dExcel.InterestRates;

using MathNet.Numerics.RootFinding;

/// <summary>
/// A collection of utility functions for dealing with interest rate curves.
/// </summary>
public static class CurveUtils
{
    /// <summary>
    /// Gets a list of all available rate indices for interest rate curve bootstrapping.
    /// </summary>
    /// <returns>A 2D column of rate index names.</returns>
    [ExcelFunction(
        Name = "d.Curve_GetRateIndices",
        Description = "Returns all available rate indices for interest rate curve bootstrapping.",
        Category = "∂Excel: Interest Rates")]
    public static object GetRateIndices()
    {
        List<string> indices = Enum.GetNames(typeof(RateIndices)).Select(x => x.ToString().ToUpper()).ToList();
        object[,] output = new object[indices.Count + 1, 1];
        output[0, 0] = "Rate Indices";
        int i = 1;
        foreach (string index in indices)
        {
            output[i++, 0] = index;
        }
        return output;
    }
   
    /// <summary>
    /// Gets all "details" stored with a curve in the DataObject controller e.g., day count convention, interpolation
    /// method, etc.
    /// </summary>
    /// <param name="handle">The 'handle' or name used to refer to the object in memory.</param>
    /// <returns>A <see cref="CurveDetails"/> object.</returns>
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
    /// Creates a QLNet YieldTermStructure curve object which is stored in the DataObjectController.
    /// </summary>
    /// <param name="handle">Handle or name to extract curve from DataObjectController.</param>
    /// <param name="curveParameters">The parameters for curve construction e.g. interpolation, day count convention etc.</param>
    /// <param name="datesRange">The dates for the corresponding discount factors.</param>
    /// <param name="discountFactorsRange">The discount factors for the corresponding dates.</param>
    /// <returns>A string containing the handle and time stamp.</returns>
    [ExcelFunction(
        Name = "d.Curve_CreateFromDiscountFactors",
        Description = 
            "Creates an interest rate curve given dates and corresponding discount factors.\n" +
            "Use 'd.Curve_GetInterpolationMethodsForDiscountFactors' to view available interpolation methods.",
        Category = "∂Excel: Interest Rates",
        IsVolatile = true)]
    public static string CreateFromDiscountFactors(
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

        if (Math.Abs((double)discountFactorsRange[0, 0] - 1.0) > 1e-10)
        {
            return CommonUtils.DExcelErrorMessage("Initial discount factor must be 1.");
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

        QL.YieldTermStructure discountCurve;
        
        if (interpolationParameter.IgnoreCaseEquals(CurveInterpolationMethods.CubicSpline_On_DiscountFactors))
        {
            discountCurve = 
                new QL.NaturalCubicDiscountCurve (
                    dates: new QL.DateVector(dates),
                    discounts: new QL.DoubleVector(discountFactors), 
                    dayCounter: dayCountConvention, 
                    calendar: calendar);
        }
        else if (interpolationParameter.IgnoreCaseEquals(CurveInterpolationMethods.Exponential_On_DiscountFactors))
        {
            discountCurve = 
                new QL.DiscountCurve(
                    dates: new QL.DateVector(dates), 
                    discounts: new QL.DoubleVector(discountFactors),
                    dayCounter: dayCountConvention, 
                    calendar: calendar);
        }
        else if (interpolationParameter.IgnoreCaseEquals(CurveInterpolationMethods.LogCubic_On_DiscountFactors))
        {
            discountCurve = 
                new QL.MonotonicLogCubicDiscountCurve (
                    dates: new QL.DateVector(dates), 
                    discounts: new QL.DoubleVector(discountFactors), 
                    dayCounter: dayCountConvention, 
                    calendar: calendar);
            
        }
        else if (interpolationParameter.IgnoreCaseEquals(CurveInterpolationMethods.NaturalLogCubic_On_DiscountFactors))
        {
            discountCurve = 
                new QL.NaturalLogCubicDiscountCurve(
                    dates: new QL.DateVector(dates), 
                    discounts: new QL.DoubleVector(discountFactors),
                    dayCounter: dayCountConvention, 
                    calendar: calendar);
        }
        else
        {
            return CommonUtils.DExcelErrorMessage($"Unsupported interpolation method: {interpolationParameter}");
        }

        bool allowExtrapolation =
            ExcelTableUtils.GetTableValue<bool?>(curveParameters, "Value", "AllowExtrapolation") ?? false;

        if (allowExtrapolation)
        {
            discountCurve.enableExtrapolation();     
        }
        
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

    /// <summary>
    /// Creates a QLNet YieldTermStructure curve object which is stored in the DataObjectController.
    /// </summary>
    /// <param name="handle">Handle or name to extract curve from DataObjectController.</param>
    /// <param name="curveParameters">The parameters for curve construction e.g. interpolation, day count convention etc.</param>
    /// <param name="datesRange">The dates for the corresponding discount factors.</param>
    /// <param name="zeroRatesRange">The discount factors for the corresponding dates.</param>
    /// <returns>A string containing the handle and time stamp.</returns>
    [ExcelFunction(
        Name = "d.Curve_CreateZeroRates",
        Description =
            "Creates an interest rate curve given dates and corresponding zero rates.\n" +
            "Use 'd.Curve_GetInterpolationMethodsForZeroRates' to view available interpolation methods.",
        Category = "∂Excel: Interest Rates",
        IsVolatile = true)]
    public static string CreateFromZeroRates(
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
            Description = "The dates for the corresponding zero rates.")]
        object[,] datesRange,
        [ExcelArgument(
            Name = "Zero Rates",
            Description = "The zero rates for the corresponding dates.")]
        object[,] zeroRatesRange)
    {
        if (datesRange.GetLength(0) != zeroRatesRange.GetLength(0))
        {
            return CommonUtils.DExcelErrorMessage("Dates and zero rates have incompatible sizes: " +
                                                  $"({datesRange.GetLength(0)} ≠ {zeroRatesRange.GetLength(0)}).");
        }

        List<QL.Date> dates = new();
        List<double> zeroRates = new();
        for (int i = 0; i < datesRange.GetLength(0); i++)
        {
            dates.Add(DateTime.FromOADate((double)datesRange[i, 0]).ToQuantLibDate());
            zeroRates.Add((double)zeroRatesRange[i, 0]);
        }

        string? dayCountConventionParameter =
            ExcelTableUtils.GetTableValue<string>(curveParameters, "Value", "DayCountConvention", 0);
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

        string? interpolationParameter =
            ExcelTableUtils.GetTableValue<string>(curveParameters, "Value", "Interpolation", 0);
        if (interpolationParameter == null)
        {
            return CommonUtils.DExcelErrorMessage("'Interpolation' not set in parameters.");
        }

        string? calendarsParameter = ExcelTableUtils.GetTableValue<string>(curveParameters, "Value", "Calendars", 0);
        IEnumerable<string>? calendars = calendarsParameter?.Split(',').Select(x => x.ToString().Trim().ToUpper());
        (QL.Calendar? calendar, string errorMessage) = DateUtils.ParseCalendars(calendarsParameter);
        
        string compoundingConventionParameter =
            ExcelTableUtils.GetTableValue<string?>(curveParameters, "Value", "CompoundingConvention", 0) ?? "";

        if (!ParserUtils.TryParseQuantLibCompoundingConvention(
                compoundingConventionParameter,
                out (QL.Compounding compounding, QL.Frequency frequency)? compoundingConvention,
                out string? compoundingConventionErrorMessage))
        {
            return compoundingConventionErrorMessage;
        }

        QL.YieldTermStructure discountCurve;

        if (interpolationParameter.IgnoreCaseEquals(CurveInterpolationMethods.Cubic_On_ZeroRates))
        {
            discountCurve =
                new QL.CubicZeroCurve(
                    dates: new QL.DateVector(dates),
                    yields: new QL.DoubleVector(zeroRates),
                    dayCounter: dayCountConvention,
                    calendar: calendar,
                    i: new QL.Cubic(),
                    compounding: compoundingConvention.Value.compounding,
                    frequency: compoundingConvention.Value.frequency);
        }
        else if (interpolationParameter.IgnoreCaseEquals(CurveInterpolationMethods.Linear_On_ZeroRates))
        {
            discountCurve =
                new QL.ZeroCurve(
                    dates: new QL.DateVector(dates),
                    yields: new QL.DoubleVector(zeroRates),
                    dayCounter: dayCountConvention,
                    calendar: calendar,
                    i: new QL.Linear(),
                    compounding: compoundingConvention.Value.compounding,
                    frequency: compoundingConvention.Value.frequency);
        }
        else if (interpolationParameter.IgnoreCaseEquals(CurveInterpolationMethods.NaturalCubic_On_ZeroRates))
        {
            discountCurve =
                new QL.NaturalCubicZeroCurve(
                    dates: new QL.DateVector(dates),
                    yields: new QL.DoubleVector(zeroRates),
                    dayCounter: dayCountConvention,
                    calendar: calendar,
                    i: new QL.SplineCubic(),
                    compounding: compoundingConvention.Value.compounding,
                    frequency: compoundingConvention.Value.frequency);
        }
        else
        {
            return CommonUtils.DExcelErrorMessage($"Unsupported interpolation method: {interpolationParameter}");
        }
        
        bool allowExtrapolation =
            ExcelTableUtils.GetTableValue<bool?>(curveParameters, "Value", "AllowExtrapolation") ?? false;

        if (allowExtrapolation)
        {
            discountCurve.enableExtrapolation();     
        }

        CurveDetails curveDetails =
            new(
                termStructure: discountCurve,
                dayCountConvention: dayCountConvention,
                interpolation: interpolationParameter,
                discountFactorDates: dates.Select(x => x.ToDateTime()),
                discountFactors: zeroRates);

        DataObjectController dataObjectController = DataObjectController.Instance;
        return dataObjectController.Add(handle, curveDetails);
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
                discountFactors[i, 0] = curve.discount(((double)dates[i]).ToQuantLibDate());
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

    /// <summary>
    /// Gets all the available interpolation methods for discount factors. This is in particularly useful for the
    /// function <see cref="CreateFromDiscountFactors"/>.
    /// </summary>
    /// <returns>A 2D column array of interpolation methods for discount factors.</returns>
    [ExcelFunction(
        Name = "d.Curve_GetInterpolationMethodsForDiscountFactors",
        Description = 
            "Gets all the available interpolation methods for discount factors.\n" +
            "This is in particularly useful for the function d.Curve_CreateFromDiscountFactors.",
        Category = "∂Excel: Interest Rates")]
    public static object GetInterpolationMethodsForDiscountFactors()
    {
        List<string> interpolationMethods = 
            Enum.GetNames(typeof(CurveInterpolationMethods))
                .Select(x => x.ToString())
                .Where(x => x.ToUpper().Contains("DISCOUNTFACTORS"))
                .ToList();
        
        object[,] output = new object[interpolationMethods.Count + 1, 1];
        output[0, 0] = "Interpolation Methods for Discount Factors";
        int i = 1;
        foreach (string interpolationMethod in interpolationMethods)
        {
            output[i++, 0] = interpolationMethod;
        }
        
        return output;
    }

    /// <summary>
    /// Returns a QuantLib term structure from an (Excel) table containing various parameters including the (string)
    /// handle to a term structure.
    /// </summary>
    /// <param name="yieldTermStructureName">The name of the term structure parameter e.g., "DiscountCurve" or
    /// "BaseCurrencyCurve" etc.</param>
    /// <param name="table">The table of parameters containing the term structure handle.</param>
    /// <param name="columnHeaderIndex">The index (in terms of row numbers, yes, row numbers) that contains the column
    /// headers.</param>
    /// <returns>A tuple containing the yield term structure, possibly null, and an error message, also possibly null.
    /// </returns>
    public static (QL.RelinkableYieldTermStructureHandle?, string? errorMessage) GetYieldTermStructure(
        string yieldTermStructureName,
        object[,] table,
        int columnHeaderIndex,
        bool allowExtrapolation)
    {
         string? termStructureHandle =
             ExcelTableUtils.GetTableValue<string?>(
                table: table, 
                columnHeader: "Value", 
                rowHeader: yieldTermStructureName, 
                rowIndexOfColumnHeaders: columnHeaderIndex);
        
        if (termStructureHandle is null)
        {
            return (null, CommonUtils.CurveParameterMissingErrorMessage(yieldTermStructureName)); 
        }
        
        QL.RelinkableYieldTermStructureHandle termStructure = new();
        QL.YieldTermStructure? tempYieldTermStructure = GetCurveObject(termStructureHandle);
        termStructure.linkTo(tempYieldTermStructure);
        if (allowExtrapolation)
        {
            termStructure.enableExtrapolation();
        }
        
        return (termStructure, null);
    }
}
