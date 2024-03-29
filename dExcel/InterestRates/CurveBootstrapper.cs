﻿using System.Diagnostics.CodeAnalysis;
using System.Text.RegularExpressions;
using dExcel.CommonEnums;
using dExcel.Dates;
using dExcel.ExcelUtils;
using dExcel.Utilities;
using ExcelDna.Integration;
using Omicron;
using QL = QuantLib;

namespace dExcel.InterestRates;

/// <summary>
/// A class containing a collection of interest rate curve bootstrapping utilities.
/// </summary>
public static class CurveBootstrapper
{
    /// <summary>
    /// Lists all available interpolation methods for interest rate curve bootstrapping.
    /// </summary>
    /// <returns>A column of all available interpolation methods for interest rate curve bootstrapping.</returns>
    [ExcelFunction(
        Name = "d.Curve_GetInterpolationMethodsForBootstrapping",
        Description = "Returns all available interpolation methods for interest rate curve bootstrapping.",
        Category = "∂Excel: Interest Rates")]
    public static object GetInterpolationMethodsForBootstrapping()
    {
        Array methods = Enum.GetValues(typeof(CurveInterpolationMethods));
        object[,] output = new object[methods.Length + 1, 1];
        output[0, 0] = "IR Bootstrapping Interpolation Methods";
        int i = 1;
        foreach (CurveInterpolationMethods method in methods)
        {
            output[i++, 0] = method.ToString();
        }
        
        return output;
    }
    
    /// <summary>
    /// Gets the IBOR index for the given name and tenor and can also apply the forecast curve if supplied.
    /// </summary>
    /// <param name="indexName">The index name.</param>
    /// <param name="indexTenor">The index tenor e.g., "3M", "1Y" etc.</param>
    /// <param name="forecastCurve">The forecast curve for the index.</param>
    /// <returns>The IBOR index, if successful, otherwise null.</returns>
    public static QL.IborIndex? GetIborIndex(
        string indexName,
        string? indexTenor,
        QL.RelinkableYieldTermStructureHandle? forecastCurve = null)
    {
        if (!Enum.TryParse(indexName.Replace("-", "_"), out RateIndices iborName))
        {
            return null;
        }
        
        QL.IborIndex? index;
        if (forecastCurve is null)
        {
            index =
                iborName switch
                {
                    RateIndices.EURIBOR => new QL.Euribor(new QL.Period(indexTenor)),
                    RateIndices.FEDFUND => new QL.FedFunds(),
                    RateIndices.JIBAR => new QL.Jibar(new QL.Period(indexTenor)),
                    RateIndices.USD_LIBOR => new QL.USDLibor(new QL.Period(indexTenor)),
                    RateIndices.SOFR => new QL.Sofr(),
                    _ => null,
                };
        }
        else 
        {
            index =
                iborName switch
                {
                    RateIndices.EURIBOR => new QL.Euribor(new QL.Period(indexTenor), forecastCurve),
                    RateIndices.FEDFUND => new QL.FedFunds(forecastCurve),
                    RateIndices.JIBAR => new QL.Jibar(new QL.Period(indexTenor), forecastCurve),
                    RateIndices.USD_LIBOR => new QL.USDLibor(new QL.Period(indexTenor), forecastCurve),
                    RateIndices.SOFR => new QL.Sofr(forecastCurve),
                    _ => null,
                };
        }

        return index;
    }

    /// <summary>
    /// Bootstraps an interest rate curve. It supports multi-curve bootstrapping.
    /// Available Indices: EURIBOR, FEDFUND (OIS), JIBAR, USD-LIBOR.
    /// </summary>
    /// <param name="handle">The 'handle' or name used to refer to the object in memory.
    /// Each object in a workbook must have a unique handle.</param>
    /// <param name="curveParameters">The parameters required to construct the curve.</param>
    /// <param name="customRateIndex">(Optional)A custom rate index.</param>
    /// <param name="instrumentGroups">The list of instrument groups used in the bootstrapping.</param>
    /// <returns>A handle to a bootstrapped curve.</returns>
    [ExcelFunction(
        Name = "d.Curve_Bootstrap",
        Description = "Bootstraps a single currency interest rate curve. Supports multi-curve bootstrapping.",
        Category = "∂Excel: Interest Rates",
        IsVolatile = true,
        IsMacroType = true)]
    public static string Bootstrap(
        [ExcelArgument(Name = "Handle", Description = DescriptionUtils.Handle)]
        string handle,
        [ExcelArgument(
            Name = "Curve Parameters",
            Description = 
                "The curves parameters: " +
                "'BaseDate', 'RateIndexName', 'RateIndexTenor', 'Interpolation', (Optional)'DiscountCurveHandle', " +
                "(Optional)'AllowExtrapolation' (Default = False)")]
        object[,] curveParameters,
        [ExcelArgument(
            Name = "(Optional)Custom Rate Index",
            Description =
                "Only populate this if you have NOT supplied a 'RateIndexName' in the curve parameters.")]
        object[,]? customRateIndex = null,
        [ExcelArgument(
            Name = "Instrument Groups",
            Description = "The instrument groups used to bootstrap the curve e.g., 'Deposits', 'FRAs', 'Swaps'.")]
        params object[] instrumentGroups)
    {
        int columnHeaderIndex = ExcelTableUtils.GetRowIndex(curveParameters, "Parameter");
        
        DateTime baseDate =
            ExcelTableUtils.GetTableValue<DateTime>(curveParameters, "Value", "Base Date", columnHeaderIndex);
        
        if (baseDate == default)
        {
            return nameof(baseDate).CurveParameterMissingErrorMessage();
        }

        QL.Settings.instance().setEvaluationDate(baseDate.ToQuantLibDate());

        string? rateIndexName =
            ExcelTableUtils.GetTableValue<string?>(curveParameters, "Value", "Rate Index Name", columnHeaderIndex);
        
        if (rateIndexName is null && customRateIndex is null)
        {
            return nameof(rateIndexName).CurveParameterMissingErrorMessage();
        }

        string? rateIndexTenor =
            ExcelTableUtils.GetTableValue<string?>(curveParameters, "Value", "Rate Index Tenor", columnHeaderIndex);
        
        if (rateIndexTenor is null && customRateIndex is null && rateIndexName != RateIndices.FEDFUND.ToString() && rateIndexName != RateIndices.SOFR.ToString())
        {
            return nameof(rateIndexTenor).CurveParameterMissingErrorMessage();
        }

        string? interpolation =
            ExcelTableUtils.GetTableValue<string?>(curveParameters, "Value", "Interpolation", columnHeaderIndex);
        
        if (interpolation == null)
        {
            return nameof(interpolation).CurveParameterMissingErrorMessage();
        }
        
        int? settlementDays =
            ExcelTableUtils.GetTableValue<int>(curveParameters, "Value", "Settlement Days", columnHeaderIndex);

        int? paymentLagDays =
            ExcelTableUtils.GetTableValue<int>(curveParameters, "Value", "Payment Lag Days", columnHeaderIndex);

        bool allowExtrapolation =
            ExcelTableUtils.GetTableValue<bool?>(curveParameters, "Value", "Allow Extrapolation", columnHeaderIndex) ??
            false;


        QL.IborIndex? rateIndex = (customRateIndex is null)? GetIborIndex(rateIndexName, rateIndexTenor, null) : null;
        if (rateIndex is null)
        {
            string? indexName = ExcelTableUtils.GetTableValue<string>(customRateIndex, "Value", "Name");
            string? tenor = ExcelTableUtils.GetTableValue<string>(customRateIndex, "Value", "Tenor");
            int? customSettlementDays = ExcelTableUtils.GetTableValue<int>(customRateIndex, "Value", "Settlement Days");
            string? customCurrencyToParse = ExcelTableUtils.GetTableValue<string>(customRateIndex, "Value", "Currency");
            ParserUtils.TryParseQuantLibCurrency(
                customCurrencyToParse, 
                out QL.Currency? customCurrency,
                out string? customCurrencyErrorMessage);

            string? calendarsToParse = ExcelTableUtils.GetTableValue<string>(customRateIndex, "Value", "Calendars");
            DateUtils.TryParseCalendars(
                calendarsToParse, 
                out QL.Calendar? customCalendar,
                out string customCalendarErrorMessage);

            string? businessDayConventionToParse =
                ExcelTableUtils.GetTableValue<string>(customRateIndex, "Value", "Business Day Convention");
            (QL.BusinessDayConvention? customBusinessDayConvention, string? customBusinessDayConventionErrorMessage) =
                DateUtils.ParseBusinessDayConvention(businessDayConventionToParse);

            string? dayCountConventionToParse =
                ExcelTableUtils.GetTableValue<string>(customRateIndex, "Value", "Day Count Convention");

            QL.DayCounter? customDayCountConvention = DateUtils.ParseDayCountConvention(dayCountConventionToParse);
            rateIndex = new QL.IborIndex(indexName, new QL.Period(tenor), (int)customSettlementDays, customCurrency,
                customCalendar, (QL.BusinessDayConvention)customBusinessDayConvention, false, customDayCountConvention);
        }

        if (rateIndex is null)
        {
            return CommonUtils.DExcelErrorMessage($"Unsupported rate index: '{rateIndexName}'");
        }

        // This is where the actual bootstrapping occurs.
        QL.RateHelperVector rateHelpers = new();

        bool instrumentsWithNaNsFound = false;
        
        foreach (object instrumentGroup in instrumentGroups)
        {
            object[,] instruments = (object[,]) instrumentGroup;
            string? instrumentType = ExcelTableUtils.GetTableLabel(instruments);
            List<string>? tenors = ExcelTableUtils.GetColumn<string>(instruments, "Tenors");
            List<string>? fraTenors = ExcelTableUtils.GetColumn<string>(instruments, "Fra Tenors");
            List<double>? rates = ExcelTableUtils.GetColumn<double>(instruments, "Rates");
            List<bool>? includeInstruments = ExcelTableUtils.GetColumn<bool>(instruments, "Include");

            if (includeInstruments is null)
            {
                continue;
            }

            int instrumentCount = includeInstruments.Count;

            if (instrumentType.IgnoreCaseEquals("Deposit", "Deposits"))
            {
                for (int i = 0; i < instrumentCount; i++)
                {
                    if (rates is null)
                    {
                        return CommonUtils.DExcelErrorMessage("Deposit rates missing.");
                    }

                    if (tenors is null)
                    {
                        return CommonUtils.DExcelErrorMessage("Deposit tenors missing");
                    }

                    instrumentsWithNaNsFound = instrumentsWithNaNsFound || double.IsNaN(rates[i]);
                    if (includeInstruments[i])
                    {
                        rateHelpers.Add(
                            new QL.DepositRateHelper(
                                rate: rates[i],
                                tenor: new QL.Period(tenors[i]),
                                fixingDays: rateIndex.fixingDays(),
                                calendar: rateIndex.fixingCalendar(),
                                convention: rateIndex.businessDayConvention(),
                                endOfMonth: rateIndex.endOfMonth(),
                                dayCounter: rateIndex.dayCounter()));
                    }
                }
            }
            else if (instrumentType.IgnoreCaseEquals("FRA", "FRAs", "Forward Rate Agreement", "Forward Rate Agreements"))
            {
                for (int i = 0; i < instrumentCount; i++)
                {
                    if (fraTenors is null)
                    {
                        return CommonUtils.DExcelErrorMessage("FRA tenors missing.");
                    }

                    if (rates is null)
                    {
                        return CommonUtils.DExcelErrorMessage("FRA rates missing.");
                    }
                    
                    instrumentsWithNaNsFound = instrumentsWithNaNsFound || double.IsNaN(rates[i]);
                    if (includeInstruments[i])
                    {
                        string fraTenorInMonths = Regex.Match(fraTenors[i], @"^\d+").Value + "M";
                        rateHelpers.Add(
                            new QL.FraRateHelper(
                                rate: new QL.QuoteHandle(new QL.SimpleQuote(rates[i])),
                                periodToStart: new QL.Period(fraTenorInMonths),
                                iborIndex: rateIndex));
                    }
                }
            }
            else if (instrumentType.IgnoreCaseEquals("IRS", "IRSs", "Interest Rate Swap", "Interest Rate Swaps"))
            {
                for (int i = 0; i < instrumentCount; i++)
                {
                    if (rates is null)
                    {
                        return CommonUtils.DExcelErrorMessage("Swap rates missing.");
                    }

                    if (tenors is null)
                    {
                        return CommonUtils.DExcelErrorMessage("Swap tenors missing");
                    }

                    instrumentsWithNaNsFound = instrumentsWithNaNsFound || double.IsNaN(rates[i]);
                    if (includeInstruments[i])
                    {
                        // Having a discount curve is only required for multi-curve bootstrapping.
                        // Hence we don't check if the curve is null here.
                        (QL.RelinkableYieldTermStructureHandle? discountCurve, string? discountCurveErrorMessage) =
                            CurveUtils.GetYieldTermStructure(
                                yieldTermStructureName: "DiscountCurve", 
                                table: curveParameters, 
                                columnHeaderIndex: columnHeaderIndex,
                                allowExtrapolation: allowExtrapolation);
                        
                        if (discountCurve is null)
                        {
                            rateHelpers.Add(
                            new QL.SwapRateHelper(
                                rate: new QL.QuoteHandle((new QL.SimpleQuote(rates[i]))),
                                tenor: new QL.Period(tenors[i]),
                                calendar: rateIndex.fixingCalendar(),
                                fixedFrequency: rateIndex.tenor().frequency(),
                                fixedConvention: rateIndex.businessDayConvention(),
                                fixedDayCount: rateIndex.dayCounter(),
                                index: rateIndex,
                                spread: new QL.QuoteHandle(new QL.SimpleQuote(0)),
                                fwdStart: new QL.Period(0, QL.TimeUnit.Months)));
                        }
                        else
                        {
                            rateHelpers.Add(
                            new QL.SwapRateHelper(
                                rate: new QL.QuoteHandle((new QL.SimpleQuote(rates[i]))),
                                tenor: new QL.Period(tenors[i]),
                                calendar: rateIndex.fixingCalendar(),
                                fixedFrequency: rateIndex.tenor().frequency(),
                                fixedConvention: rateIndex.businessDayConvention(),
                                fixedDayCount: rateIndex.dayCounter(),
                                index: rateIndex,
                                spread: new QL.QuoteHandle(new QL.SimpleQuote(0)),
                                fwdStart: new QL.Period(0, QL.TimeUnit.Months),
                                discountingCurve: discountCurve));
                        }
                    }
                }
            }
            else if (instrumentType.IgnoreCaseEquals("OIS", "OISs", "OIS Swap", "OIS Swaps", "Overnight Index Swap", "Overnight Index Swaps"))
            {
                if (rates is null)
                {
                    return CommonUtils.DExcelErrorMessage("OIS rates missing.");
                }
                
                instrumentsWithNaNsFound = instrumentsWithNaNsFound || double.IsNaN(rates[0]);
                for (int i = 0; i < instrumentCount; i++)
                {
                    if (includeInstruments[i])
                    {
                        rateHelpers.Add(
                            new QL.OISRateHelper(
                                settlementDays: (uint)settlementDays, 
                                tenor: new QL.Period(tenors?[i]),
                                rate: new QL.QuoteHandle(new QL.SimpleQuote(rates[i])), 
                                index: rateIndex as QL.OvernightIndex));
                    }
                }
            }
            else
            {
                return CommonUtils.DExcelErrorMessage($"Unknown instrument type: '{instrumentType}'");
            }
        }

        QL.YieldTermStructure? termStructure =
            BootstrapCurveFromRateHelpers(
                rateHelpers: rateHelpers, 
                referenceDate: baseDate, 
                dayCountConvention: rateIndex.dayCounter(), 
                interpolation: interpolation);
        
        if (termStructure is null)
        {
            return CommonUtils.DExcelErrorMessage($"Unknown interpolation method: '{interpolation}'");
        }
        
        if (allowExtrapolation)
        {
            termStructure.enableExtrapolation();
        }
        
        CurveDetails curveDetails = 
            new(termStructure: termStructure, 
                dayCountConvention: rateIndex.dayCounter(), 
                interpolation: interpolation, 
                discountFactorDates: null,  
                discountFactors: null, 
                instrumentGroups: instrumentGroups);
        
        DataObjectController dataObjectController = DataObjectController.Instance;
        
        string warningMessage = instrumentsWithNaNsFound ? "Instruments with NaNs found" : "";
        return dataObjectController.Add(handle, curveDetails, warningMessage);
    }

    /// <summary>
    /// Bootstraps a tenor basis, single currency interest rate curve. 
    /// Available Indices: EURIBOR, FEDFUND (OIS), JIBAR, USD-LIBOR, SOFR.
    /// </summary>
    /// <param name="handle">The 'handle' or name used to refer to the object in memory.
    /// Each object in a workbook must have a unique handle.</param>
    /// <param name="curveParameters">The parameters required to construct the curve.</param>
    /// <param name="customBaseIndex">(Optional)A custom base rate index.</param>
    /// <param name="customSpreadIndex">(Optional)A custom spread rate index.</param>
    /// <param name="instrumentGroups">The list of instrument groups used in the bootstrapping.</param>
    /// <returns>A handle to a bootstrapped curve.</returns>
    [ExcelFunction(
        Name = "d.Curve_BootstrapTenorBasisCurve",
        Description = "Bootstraps a single currency interest rate curve. Supports multi-curve bootstrapping.",
        Category = "∂Excel: Interest Rates",
        IsVolatile = true,
        IsMacroType = true)]
    public static string BootstrapTenorBasisCurve(
        [ExcelArgument(Name = "Handle", Description = DescriptionUtils.Handle)]
        string handle,
        [ExcelArgument(
            Name = "Curve Parameters",
            Description = "The curves parameters.")]
        object[,] curveParameters,
        [ExcelArgument(
            Name = "(Optional)Custom Base Index",
            Description =
                "Only populate this if you have NOT supplied a 'RateIndexName' in the curve parameters.")]
        object[,]? customBaseIndex = null,
        [ExcelArgument(
            Name = "(Optional)Custom Spread Index",
            Description =
                "Only populate this if you have NOT supplied a 'RateIndexName' in the curve parameters.")]
        object[,]? customSpreadIndex = null,
        [ExcelArgument(
            Name = "Instrument Groups",
            Description = "The instrument groups used to bootstrap the curve e.g., 'Deposits', 'FRAs', 'Swaps'.")]
        params object[] instrumentGroups)
    {
        int columnHeaderIndex = ExcelTableUtils.GetRowIndex(curveParameters, "Parameter");
        
        DateTime baseDate =
            ExcelTableUtils.GetTableValue<DateTime>(curveParameters, "Value", "BaseDate", columnHeaderIndex);
        
        if (baseDate == default)
        {
            return nameof(baseDate).CurveParameterMissingErrorMessage();
        }

        QL.Settings.instance().setEvaluationDate(baseDate.ToQuantLibDate());

        bool allowExtrapolation =
            ExcelTableUtils.GetTableValue<bool?>(curveParameters, "Value", "Allow Extrapolation", columnHeaderIndex) ??
            false;

        string? baseIndexName =
            ExcelTableUtils.GetTableValue<string?>(curveParameters, "Value", "Base Index Name", columnHeaderIndex);
        
        if (baseIndexName is null && customBaseIndex is null)
        {
            return nameof(baseIndexName).CurveParameterMissingErrorMessage();
        }

        string? baseIndexTenor =
            ExcelTableUtils.GetTableValue<string?>(curveParameters, "Value", "Base Index Tenor", columnHeaderIndex);
        
        if (baseIndexTenor is null && customBaseIndex is null && baseIndexName != RateIndices.FEDFUND.ToString())
        {
            return nameof(baseIndexTenor).CurveParameterMissingErrorMessage();
        }

        string? baseIndexDiscountCurveHandle =
            ExcelTableUtils.GetTableValue<string?>(curveParameters, "Value", "Base Index Discount Curve", columnHeaderIndex);

        string? spreadIndexName =
            ExcelTableUtils.GetTableValue<string?>(curveParameters, "Value", "Spread Index", columnHeaderIndex);

        string? spreadIndexTenor =
            ExcelTableUtils.GetTableValue<string?>(curveParameters, "Value", "Spread Index Tenor", columnHeaderIndex);

        string? interpolation =
            ExcelTableUtils.GetTableValue<string?>(curveParameters, "Value", "Interpolation", columnHeaderIndex);
        
        if (interpolation == null)
        {
            return nameof(interpolation).CurveParameterMissingErrorMessage();
        }

        (QL.RelinkableYieldTermStructureHandle? baseIndexForecastCurve, string? baseIndexForecastCurveErrorMessage) =
            CurveUtils.GetYieldTermStructure(
                yieldTermStructureName: "BaseIndexForecastCurve", 
                table: curveParameters, 
                columnHeaderIndex: columnHeaderIndex,
                allowExtrapolation: allowExtrapolation);
        
        if (baseIndexForecastCurveErrorMessage is not null)
        {
            return baseIndexForecastCurveErrorMessage;
        }
        
        (QL.RelinkableYieldTermStructureHandle? baseIndexDiscountCurve, string? baseIndexDiscountCurveErrorMessage) =
            CurveUtils.GetYieldTermStructure(
                yieldTermStructureName: "BaseIndexDiscountCurve", 
                table: curveParameters, 
                columnHeaderIndex: columnHeaderIndex,
                allowExtrapolation: allowExtrapolation);
            
        if (baseIndexDiscountCurveErrorMessage is not null)
        {
            return baseIndexDiscountCurveErrorMessage;
        }
        
        QL.IborIndex? baseIndex = GetIborIndex(baseIndexName, baseIndexTenor, baseIndexForecastCurve);
        QL.IborIndex? otherIndex = GetIborIndex(spreadIndexName, spreadIndexTenor,null);

        if (baseIndex is null)
        {
            return CommonUtils.DExcelErrorMessage($"Unsupported rate index: '{baseIndexName}'");
        }

        QL.RateHelperVector rateHelpers = new();
        bool instrumentsWithNaNsFound = false;

        foreach (object instrumentGroup in instrumentGroups)
        {
            object[,] instruments = (object[,]) instrumentGroup;
            string? instrumentType = ExcelTableUtils.GetTableLabel(instruments);
            List<string>? tenors = ExcelTableUtils.GetColumn<string>(instruments, "Tenors");
            List<double>? basisSpreads = ExcelTableUtils.GetColumn<double>(instruments, "Basis Spreads");
            List<bool>? includeInstruments = ExcelTableUtils.GetColumn<bool>(instruments, "Include");

            if (includeInstruments is null)
            {
                continue;
            }

            int instrumentCount = includeInstruments.Count;

            if (instrumentType.IgnoreCaseEquals("Basis Swap", "Basis Swaps", "Tenor Basis Swap", "Tenor Basis Swaps"))
            {
                if (basisSpreads is null)
                {
                    return CommonUtils.DExcelErrorMessage("Basis spreads missing.");
                }

                if (spreadIndexName is null)
                {

                }
                
                for (int i = 0; i < instrumentCount; i++)
                {
                    instrumentsWithNaNsFound = instrumentsWithNaNsFound || double.IsNaN(basisSpreads[i]);
                    
                    if (includeInstruments[i])
                    {
                        rateHelpers.Add(
                            new QL.IborIborBasisSwapRateHelper(
                                settlementDays: 2, 
                                tenor: new QL.Period(tenors?[i]),
                                basis: new QL.QuoteHandle(new QL.SimpleQuote(basisSpreads[i])),
                                baseIndex: baseIndex,
                                calendar: baseIndex.fixingCalendar(),
                                convention: baseIndex.businessDayConvention(),
                                endOfMonth: false,
                                otherIndex: otherIndex,
                                discountHandle: baseIndexDiscountCurve,
                                bootstrapBaseCurve: false));
                    }
                }
            }
        }

        QL.YieldTermStructure? termStructure =
            BootstrapCurveFromRateHelpers(
                rateHelpers: rateHelpers, 
                referenceDate: baseDate, 
                dayCountConvention: baseIndex.dayCounter(), 
                interpolation: interpolation);
        
        if (termStructure is null)
        {
            return CommonUtils.DExcelErrorMessage($"Unknown interpolation method: '{interpolation}'");
        }
        
        if (allowExtrapolation)
        {
            termStructure.enableExtrapolation();
        }
        
        CurveDetails curveDetails = 
            new(termStructure: termStructure, 
                dayCountConvention: baseIndex.dayCounter(), 
                interpolation: interpolation, 
                discountFactorDates: null,  
                discountFactors: null, 
                instrumentGroups: instrumentGroups);
        
        DataObjectController dataObjectController = DataObjectController.Instance;
        string warningMessage = instrumentsWithNaNsFound ? "Instruments with NaNs found" : "";
        return dataObjectController.Add(handle, curveDetails, warningMessage);
    }
    
    /// <summary>
    /// Bootstraps a curve (single or multi-curve) given a list of rate helpers.
    /// </summary>
    /// <param name="rateHelpers">Rate helpers.</param>
    /// <param name="referenceDate">The curve base date.</param>
    /// <param name="dayCountConvention">The day count convention.</param>
    /// <param name="interpolation">The interpolation method.</param>
    /// <returns>A bootstrapped term structure, if it succeeds, otherwise null.</returns>
    public static QL.YieldTermStructure? BootstrapCurveFromRateHelpers(
        QL.RateHelperVector rateHelpers,
        DateTime referenceDate,
        QL.DayCounter dayCountConvention,
        string interpolation)
    {
        QL.Date curveDate = referenceDate.ToQuantLibDate();
        if (interpolation.IgnoreCaseEquals(CurveInterpolationMethods.Flat_On_ForwardRates))
        {
            return new QL.PiecewiseFlatForward(curveDate, rateHelpers, dayCountConvention);
        }

        if (interpolation.IgnoreCaseEquals(CurveInterpolationMethods.CubicSpline_On_DiscountFactors))
        {
            return new QL.PiecewiseSplineCubicDiscount(curveDate, rateHelpers, dayCountConvention);
        }

        if (interpolation.IgnoreCaseEquals(CurveInterpolationMethods.Exponential_On_DiscountFactors))
        {
            return new QL.PiecewiseLogLinearDiscount(curveDate, rateHelpers, dayCountConvention);
        }
        
        if (interpolation.IgnoreCaseEquals(CurveInterpolationMethods.LogCubic_On_DiscountFactors))
        {
            return new QL.PiecewiseLogCubicDiscount(curveDate, rateHelpers, dayCountConvention);
        }
    
        if (interpolation.IgnoreCaseEquals(CurveInterpolationMethods.NaturalLogCubic_On_DiscountFactors))
        {
            return new QL.PiecewiseNaturalLogCubicDiscount(curveDate, rateHelpers, dayCountConvention);
        }

        if (interpolation.IgnoreCaseEquals(CurveInterpolationMethods.Cubic_On_ZeroRates))
        {
            return new QL.PiecewiseCubicZero(curveDate, rateHelpers, dayCountConvention);
        }
        
        if (interpolation.IgnoreCaseEquals(CurveInterpolationMethods.Linear_On_ZeroRates))
        {
            return new QL.PiecewiseLinearZero(curveDate, rateHelpers, dayCountConvention);
        }
        
        if (interpolation.IgnoreCaseEquals(CurveInterpolationMethods.NaturalCubic_On_ZeroRates))
        {
            return new QL.PiecewiseNaturalCubicZero(curveDate, rateHelpers, dayCountConvention);
        }
        
        return null;
    }
    
    /// <summary>
    /// Extracts and bootstraps a curve from the Omicron database.
    /// </summary>
    /// <param name="handle">The 'handle' or name used to refer to the object in memory.
    /// Each object in a workbook must have a a unique handle.</param>
    /// <param name="curveName">The name of the curve in Omicron. Current available options are:
    /// 'ZAR-Swap', 'USD-OIS'</param>
    /// <param name="baseDate"></param>
    /// <param name="interpolation"></param>
    /// <returns></returns>
    [ExcelFunction(
        Name = "d.Curve_Get",
        Description = "Extracts and bootstraps a curve from the Omicron database.",
        Category = "∂Excel: Interest Rates")]
    public static string Get(
        [ExcelArgument(Name = "Handle", Description = DescriptionUtils.Handle)]
        string handle,
        [ExcelArgument(Name = "Curve Name", Description = "The name of the curve in Omicron.")]
        string curveName,
        [ExcelArgument(
            Name = "Base Date",
            Description = "The base date of the curve i.e., the date for which to extract the curve.")]
        DateTime baseDate,
        [ExcelArgument(
            Name = "DF Interpolation",
            Description = 
                "(Optional)The discount factor interpolation style.\n" +
                "Default = 'Exponential_On_DiscountFactors'.")]
        string interpolation = "")
    {
        interpolation = 
            interpolation != "" ? interpolation : CurveInterpolationMethods.Exponential_On_DiscountFactors.ToString();

        if (!TryParseCurveNameToRateIndex(
                curveName, 
                out (string name, string tenor)? index,
                out string curveNameErrorMessage))
        {
            return curveNameErrorMessage; 
        }
            
        List<QuoteValue> quoteValues;
        try
        {
            // quoteValues =
            //     (List<QuoteValue>)ExcelAsyncUtil.Run(
            //         nameof(Get), 
            //         new object[] { index, baseDate },
            //         () => OmicronUtils.OmicronUtils.GetSwapCurveQuotes(
            //             index: index.Value.name.Replace("_", "-"), 
            //             spreadIndex: null,
            //             quotes: null,
            //             requisitionId: 1, 
            //             marketDataDate: baseDate.ToString("yyyy-MM-dd")));

            quoteValues =
                Task.Run(
                    function: () => OmicronUtils.OmicronUtils.GetAllSwapCurveQuotes(
                        index: index.Value.name, 
                        tenor: new QL.Period(index.Value.tenor).ToOmicronTenor(), 
                        marketDataDate: baseDate))
                    .Result;
        }
        catch (Exception ex)
        {
            if (!NetworkUtils.GetVpnConnectionStatus())
            {
                return CommonUtils.DExcelErrorMessage("Not connected to Deloitte network/VPN.");
            }

            return CommonUtils.DExcelErrorMessage($"Unknown error: {ex.Message}");
        }

        object[,] curveParameters =
        {
            {"Curve Utils Parameters", ""},
            {"Parameter", "Value"},
            {"Base Date", baseDate.ToOADate()},
            {"Rate Index Name", index.Value.name},
            {"Rate Index Tenor", index.Value.tenor},
            {"Interpolation", interpolation},
        };

        List<QuoteValue> deposits = quoteValues.Where(x => x.Type.GetType() == typeof(RateIndex)).ToList();
        deposits = deposits.OrderBy(x => ((RateIndex)x.Type).Tenor, new TenorComparer()).ToList();
        object[,] depositInstruments = new object[deposits.Count + 2, 4];
        depositInstruments[0, 0] = "Deposits";
        depositInstruments[1, 0] = "Tenors";
        depositInstruments[1, 1] = "Rate Index";
        depositInstruments[1, 2] = "Rates";
        depositInstruments[1, 3] = "Include";

        int row = 2;
        foreach (QuoteValue deposit in deposits)
        {
            depositInstruments[row, 0] = ((RateIndex) deposit.Type).Tenor.ToString();
            depositInstruments[row, 1] = ((RateIndex) deposit.Type).Name;
            depositInstruments[row, 2] = double.IsNaN(deposit.Value) ? "NaN" : deposit.Value;
            depositInstruments[row, 3] = double.IsNaN(deposit.Value) ? "FALSE" : "TRUE";
            row++;
        }

        List<QuoteValue> fras = quoteValues.Where(x => x.Type.GetType() == typeof(Fra)).ToList();
        fras = fras.OrderBy(x => ((Fra)x.Type).Tenor, new TenorComparer()).ToList();
        object[,] fraInstruments = new object[fras.Count + 2, 4];
        row = 2;
        fraInstruments[0, 0] = "FRAs";
        fraInstruments[1, 0] = "Fra Tenors";
        fraInstruments[1, 1] = "Rate Index";
        fraInstruments[1, 2] = "Rates";
        fraInstruments[1, 3] = "Include";

        foreach (QuoteValue fra in fras)
        {
            // TODO: Ensure the amount is always in months.
            fraInstruments[row, 0] = $"{((Fra) fra.Type).Tenor.Amount}x{((Fra) fra.Type).Tenor.Amount + 3}";
            fraInstruments[row, 1] = ((Fra) fra.Type).ReferenceIndex.Name;
            fraInstruments[row, 2] = double.IsNaN(fra.Value) ? "NaN" : fra.Value;
            fraInstruments[row, 3] = double.IsNaN(fra.Value) ? "FALSE" : "TRUE";
            row++;
        }

        List<QuoteValue> swaps = quoteValues.Where(x => x.Type.GetType() == typeof(InterestRateSwap)).ToList();
        swaps = swaps.OrderBy(x => ((InterestRateSwap)x.Type).Tenor, new TenorComparer()).ToList();
        object[,] swapInstruments = new object[swaps.Count + 2, 4];
        swapInstruments[0, 0] = "Interest Rate Swaps";
        swapInstruments[1, 0] = "Tenors";
        swapInstruments[1, 1] = "Rate Index";
        swapInstruments[1, 2] = "Rates";
        swapInstruments[1, 3] = "Include";

        row = 2;
        foreach (QuoteValue swap in swaps)
        {
            swapInstruments[row, 0] = ((InterestRateSwap) swap.Type).Tenor.ToString();
            swapInstruments[row, 1] = ((InterestRateSwap) swap.Type).ReferenceIndex.Name;
            swapInstruments[row, 2] = double.IsNaN(swap.Value) ? "NaN" : swap.Value;
            swapInstruments[row, 3] = double.IsNaN(swap.Value) ? "FALSE" : "TRUE";
            row++;
        }
        
        List<QuoteValue> overnightIndexSwaps = quoteValues.Where(x => x.Type.GetType() == typeof(Ois)).ToList();
        overnightIndexSwaps = overnightIndexSwaps.OrderBy(x => ((Ois)x.Type).Tenor, new TenorComparer()).ToList();
        object[,] oisInstruments = new object[overnightIndexSwaps.Count + 2, 3];
        oisInstruments[0, 0] = "OISs";
        oisInstruments[1, 0] = "Tenors";
        oisInstruments[1, 1] = "Rates";
        oisInstruments[1, 2] = "Include";

        row = 2;
        foreach (QuoteValue ois in overnightIndexSwaps)
        {
            oisInstruments[row, 0] = ((Ois)ois.Type).Tenor.ToString();
            oisInstruments[row, 1] = double.IsNaN(ois.Value) ? "NaN" : ois.Value;
            oisInstruments[row, 2] = double.IsNaN(ois.Value) ? "FALSE" : "TRUE";
            row++;
        }

        List<object> instruments = new();
        if (deposits.Any()) instruments.Add(depositInstruments);
        if (fras.Any()) instruments.Add(fraInstruments);
        if (swaps.Any()) instruments.Add(swapInstruments);
        if (overnightIndexSwaps.Any()) instruments.Add(oisInstruments);

        return Bootstrap(handle, curveParameters, null, instruments.ToArray());
    }

    /// <summary>
    /// Tries to parse a curve name to a rate index pair of "(name, tenor)" e.g., "ZAR_Swap" => ("JIBAR", "3M").
    /// </summary>
    /// <param name="curveName">The curve name e.g., "ZAR_Swap", "USD_Libor"</param>
    /// <param name="rateIndex">A tuple consisting of the rate index name and tenor.</param>
    /// <param name="errorMessage">An error message if it can't parse the curve name.</param>
    /// <returns>True if it can parse the curve name, otherwise false.</returns>
    public static bool TryParseCurveNameToRateIndex(
        string curveName, 
        [NotNullWhen(true)] 
        out (string name, string tenor)? rateIndex, 
        out string errorMessage)
    {
        errorMessage = "";
        curveName = curveName.Replace("-", "_");
        if (curveName.IgnoreCaseEquals(OmicronSwapCurves.ZAR_Swap.ToString()))
        {
            rateIndex = (RateIndices.JIBAR.ToString(), "3M");
            return true;
        }
        
        if (curveName.IgnoreCaseEquals(OmicronSwapCurves.USD_Swap.ToString()))
        {
            rateIndex = (RateIndices.USD_LIBOR.ToString(), "3M");
            return true;
        }
        
        if (curveName.IgnoreCaseEquals(OmicronSwapCurves.USD_OIS.ToString()))
        {
            rateIndex = (RateIndices.FEDFUND.ToString(), "1D");
            return true;
        }

        if (curveName.IgnoreCaseEquals(OmicronSwapCurves.SOFR.ToString()))
        {
            rateIndex = (RateIndices.SOFR.ToString(), "1D");
            return true;
        }

        rateIndex = null;
        errorMessage = CommonUtils.DExcelErrorMessage($"Unsupported curve name: {curveName}");
        return false;
    }
   
    /// <summary>
    /// Extracts the underlying swap curve quotes for a given curve.
    /// </summary>
    /// <param name="curveName">The curve name e.g., "USD-Swap", "ZAR-Swap", etc.</param>
    /// <param name="baseDate">The base date of the curve.</param>
    /// <returns>A 2D array of quotes for the curve.</returns>
    public static object GetSwapCurveQuotes(string curveName, DateTime baseDate)
    {
        if (!TryParseCurveNameToRateIndex(curveName, out (string name, string tenor)? index, out string errorMessage))
        {
            return errorMessage;        
        }
        
        List<QuoteValue> quoteValues = 
            (List<QuoteValue>)ExcelAsyncUtil.Run(
                nameof(GetSwapCurveQuotes),
                new object[] {index, baseDate},
                () => dExcel.OmicronUtils.OmicronUtils.GetSwapCurveQuotes(
                    index: index.Value.name.Replace("_", "-"),
                    spreadIndex: null,
                    quotes: null, 
                    requisitionId: 1, 
                    marketDataDate: baseDate.ToString("yyyy-MM-dd")));
        
        object[,] output = new object[quoteValues.Count, 1];
        for (int i = 0; i < quoteValues.Count; i++)
        {
            output[i, 0] = Regex.Replace(quoteValues[i].ToString(), @"(\d{2})/(\d{2})/(\d{4})", "$3-$2-$1");
        }   
        
        return output;
    }
    
    
    /// <summary>
    /// Extracts the underlying swap curve quotes for a given curve.
    /// </summary>
    /// <param name="curveName">The curve name e.g., "USD-Swap", "ZAR-Swap", etc.</param>
    /// <param name="baseDate">The base date of the curve.</param>
    /// <returns>A 2D array of quotes for the curve.</returns>
    [ExcelFunction(
        Name = "d.Curve_GetSwapCurveQuotes",
        Description = "Extracts the underlying swap quotes for a given curve.",
        Category = "∂Excel: Interest Rates")]
    public static object GetAllSwapCurveQuotes(string curveName, DateTime baseDate)
    {
        if (!TryParseCurveNameToRateIndex(curveName, out (string name, string tenor)? index, out string errorMessage))
        {
            return errorMessage;        
        }
        
        QL.Period tenor = new(index.Value.tenor); 
        List<QuoteValue> quoteValues = 
           Task.Run(() => OmicronUtils.OmicronUtils.GetAllSwapCurveQuotes(index.Value.name, tenor.ToOmicronTenor(), baseDate)).Result;
        
        object[,] output = new object[quoteValues.Count, 1];
        for (int i = 0; i < quoteValues.Count; i++)
        {
            output[i, 0] = Regex.Replace(quoteValues[i].ToString(), @"(\d{2})/(\d{2})/(\d{4})", "$3-$2-$1");
        }   
        
        return output;
    }
    
    /// <summary>
    /// Creates flat curve for testing purposes. Assumes simple interest rate compounding convention.
    /// </summary>
    /// <param name="handle">The 'handle' or name used to refer to the object in memory.
    /// Each object in a workbook must have a unique handle.</param>
    /// <param name="baseDate">Date used as starting point of the curve.</param>
    /// <param name="rate">Flat interest rate assuming simple compounding.</param>
    /// <param name="dayCountConventionToParse">Day count convention to be applied.</param>
    /// <returns>A handle to a bootstrapped curve.</returns>
    [ExcelFunction(
        Name = "d.Curve_CreateFlatCurve",
        Description = "Creates flat curve for testing purposes. Assumes simple interest rate compounding convention.",
        Category = "∂Excel: Interest Rates")]
    public static string CreateFlatCurve(
        [ExcelArgument(Name = "Handle", Description = DescriptionUtils.Handle)]
        string handle,
        [ExcelArgument(Name = "baseDate", Description = "The start date of the curve")]
        DateTime baseDate,
        [ExcelArgument(Name = "rate", Description = "The flat interest rate. Simple compounding convention.")]
        double rate,
        [ExcelArgument(Name = "Day count convention", Description = "The flat interest rate. Simple compounding convention.")]
        string dayCountConventionToParse)
    {

        QL.DayCounter? dayCountConvention = DateUtils.ParseDayCountConvention(dayCountConventionToParse);

        QL.FlatForward interestRateCurve =
            new(baseDate.ToQuantLibDate(),
            new QL.QuoteHandle(new QL.SimpleQuote(rate)), dayCountConvention);

        CurveDetails curveDetails = new(interestRateCurve, null, null, null, null, null);
        DataObjectController dataObjectController = DataObjectController.Instance;
        return dataObjectController.Add(handle, curveDetails);
    }
}
