using dExcel.Dates;
using dExcel.ExcelUtils;
using dExcel.InterestRates;
using dExcel.Utilities;
using ExcelDna.Integration;
using QL = QuantLib;

namespace dExcel.FX;

/// <summary>
/// A class containing a collection of FX curve bootstrapping utilities.
/// </summary>
public static class CurveBootstrapper
{
    /// <summary>
    /// Bootstraps an FX basis adjusted curve. 
    /// Available Indices: EURIBOR, FEDFUND (OIS), JIBAR, USD-LIBOR.
    /// </summary>
    /// <param name="handle">The 'handle' or name used to refer to the object in memory.
    /// Each object in a workbook must have a unique handle.</param>
    /// <param name="curveParameters">The parameters required to construct the curve.</param>
    /// <param name="customBaseCurrencyIndex">(Optional)A custom rate index for the base currency.</param>
    /// <param name="customQuoteCurrencyIndex">(Optional)A custom rate index for the quote currency.</param>
    /// <param name="instrumentGroups">The list of instrument groups used in the bootstrapping.</param>
    /// <returns>A handle to a bootstrapped curve.</returns>
    [ExcelFunction(
        Name = "d.Curve_BootstrapFxBasisCurve",
        Description = "Bootstraps an FX basis adjusted curve. Supports multi-curve bootstrapping.",
        Category = "∂Excel: Interest Rates",
        IsVolatile = true,
        IsMacroType = true)]
    public static string BootstrapFxBasisCurve(
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
            Name = "(Optional)Custom Base Currency Index",
            Description =
                "Only populate this if you have NOT supplied a 'BaseCurrencyIndex' in the curve parameters.")]
        object[,]? customBaseCurrencyIndex = null,
        [ExcelArgument(
            Name = "(Optional)Custom Quote Currency Index",
            Description =
                "Only populate this if you have NOT supplied a 'QuoteCurrencyIndex' in the curve parameters.")]
        object[,]? customQuoteCurrencyIndex = null,
        [ExcelArgument(
            Name = "Instrument Groups",
            Description = "The instrument groups used to bootstrap the curve e.g., 'FECs', 'Cross Currency Swaps'.")]
        params object[] instrumentGroups)
    {
        int columnHeaderIndex = ExcelTableUtils.GetRowIndex(curveParameters, "Parameter");
        
        DateTime baseDate =
            ExcelTableUtils.GetTableValue<DateTime>(curveParameters, "Value", "BaseDate", columnHeaderIndex);
        
        if (baseDate == default)
        {
            return CommonUtils.CurveParameterMissingErrorMessage(nameof(baseDate).ToUpper());
        }

        QL.Settings.instance().setEvaluationDate(baseDate.ToQuantLibDate());

        string? baseCurrencyIndexName =
            ExcelTableUtils.GetTableValue<string?>(curveParameters, "Value", "BaseCurrencyIndexName", columnHeaderIndex);
        
        if (baseCurrencyIndexName is null && customBaseCurrencyIndex is null)
        {
            return CommonUtils.CurveParameterMissingErrorMessage(nameof(baseCurrencyIndexName).ToUpper());
        }
        
        string? quoteCurrencyIndexName =
            ExcelTableUtils.GetTableValue<string?>(curveParameters, "Value", "QuoteCurrencyIndexName", columnHeaderIndex);
        
        if (quoteCurrencyIndexName is null && customBaseCurrencyIndex is null)
        {
            return CommonUtils.CurveParameterMissingErrorMessage(nameof(quoteCurrencyIndexName).ToUpper());
        }

        string? baseCurrencyIndexTenor =
            ExcelTableUtils.GetTableValue<string?>(curveParameters, "Value", "BaseCurrencyIndexTenor", columnHeaderIndex);
        
        if (baseCurrencyIndexTenor is null && customBaseCurrencyIndex is null && quoteCurrencyIndexName != "FEDFUND")
        {
            return CommonUtils.CurveParameterMissingErrorMessage(nameof(baseCurrencyIndexTenor).ToUpper());
        }

        string? quoteCurrencyIndexTenor =
            ExcelTableUtils.GetTableValue<string?>(curveParameters, "Value", "QuoteCurrencyIndexTenor", columnHeaderIndex);
        
        if (quoteCurrencyIndexTenor is null && customBaseCurrencyIndex is null && quoteCurrencyIndexName != "FEDFUND")
        {
            return CommonUtils.CurveParameterMissingErrorMessage(nameof(quoteCurrencyIndexTenor).ToUpper());
        }
        
        string? interpolation =
            ExcelTableUtils.GetTableValue<string?>(curveParameters, "Value", "Interpolation", columnHeaderIndex);
        
        if (interpolation == null)
        {
            return CommonUtils.CurveParameterMissingErrorMessage(nameof(interpolation).ToUpper());
        }
        
        double? spotFx =
            ExcelTableUtils.GetTableValue<double?>(curveParameters, "Value", "SpotFX", columnHeaderIndex);
        
        if (spotFx == null)
        {
            return CommonUtils.CurveParameterMissingErrorMessage(nameof(spotFx).ToUpper());
        }
        
        string? baseCurrencyDiscountCurveHandle =
            ExcelTableUtils.GetTableValue<string?>(
                table: curveParameters, 
                columnHeader: "Value", 
                rowHeader: "BaseCurrencyDiscountCurveHandle", 
                rowIndexOfColumnHeaders: columnHeaderIndex);
        
        QL.RelinkableYieldTermStructureHandle baseCurrencyDiscountCurve = new();
        if (baseCurrencyDiscountCurveHandle != null)
        {
            QL.YieldTermStructure? yieldTermStructure = CurveUtils.GetCurveObject(baseCurrencyDiscountCurveHandle);
            baseCurrencyDiscountCurve.linkTo(yieldTermStructure);
        }
        
        string? baseCurrencyForecastCurveHandle =
            ExcelTableUtils.GetTableValue<string?>(
                table: curveParameters, 
                columnHeader: "Value", 
                rowHeader: "BaseCurrencyForecastCurveHandle", 
                rowIndexOfColumnHeaders: columnHeaderIndex);
        
        QL.RelinkableYieldTermStructureHandle baseCurrencyForecastCurve = new();
        if (baseCurrencyForecastCurveHandle != null)
        {
            QL.YieldTermStructure? yieldTermStructure = CurveUtils.GetCurveObject(baseCurrencyForecastCurveHandle);
            baseCurrencyForecastCurve.linkTo(yieldTermStructure);
        }
        
        string? quoteCurrencyForecastCurveHandle =
            ExcelTableUtils.GetTableValue<string?>(
                table: curveParameters, 
                columnHeader: "Value", 
                rowHeader: "QuoteCurrencyForecastCurveHandle", 
                rowIndexOfColumnHeaders: columnHeaderIndex);
        
        QL.RelinkableYieldTermStructureHandle quoteCurrencyForecastCurve = new();
        if (baseCurrencyForecastCurveHandle != null)
        {
            QL.YieldTermStructure? yieldTermStructure = CurveUtils.GetCurveObject(quoteCurrencyForecastCurveHandle);
            quoteCurrencyForecastCurve.linkTo(yieldTermStructure);
        }
        
        bool? allowExtrapolation =
            ExcelTableUtils.GetTableValue<bool?>(curveParameters, "Value", "AllowExtrapolation", columnHeaderIndex);
        
        if (allowExtrapolation == null)
        {
            allowExtrapolation = false;
        }
        
        QL.IborIndex? quoteCurrencyIndex = 
            dExcel.InterestRates.CurveBootstrapper.GetIborIndex(
                indexName: quoteCurrencyIndexName, 
                indexTenor: quoteCurrencyIndexTenor, 
                forecastCurve: quoteCurrencyForecastCurve);
        
        QL.IborIndex? baseCurrencyIndex = 
            dExcel.InterestRates.CurveBootstrapper.GetIborIndex(
                indexName: baseCurrencyIndexName, 
                indexTenor: baseCurrencyIndexTenor, 
                forecastCurve: baseCurrencyForecastCurve);

        if (baseCurrencyIndex is null)
        {
            return CommonUtils.DExcelErrorMessage($"Unsupported rate index: {baseCurrencyIndex}");
        }

        if (quoteCurrencyIndex is null)
        {
            return CommonUtils.DExcelErrorMessage($"Unsupported rate index: {quoteCurrencyIndexName}");
        }
        
        QL.RateHelperVector rateHelpers = new();

        foreach (object instrumentGroup in instrumentGroups)
        {
            object[,] instruments = (object[,]) instrumentGroup;
            string? instrumentType = ExcelTableUtils.GetTableLabel(instruments);
            if (instrumentType is null)
            {
                return CommonUtils.DExcelErrorMessage("No instrument type found.");
            }
            
            List<string>? tenors = ExcelTableUtils.GetColumn<string>(instruments, "Tenors");
            List<double>? basisSpreads = ExcelTableUtils.GetColumn<double>(instruments, "BasisSpreads");
            List<double>? forwardPoints = ExcelTableUtils.GetColumn<double>(instruments, "ForwardPoints");
            List<int>? fixingDays = ExcelTableUtils.GetColumn<int>(instruments, "FixingDays");
            List<bool>? includeInstruments = ExcelTableUtils.GetColumn<bool>(instruments, "Include");

            if (includeInstruments is null)
            {
                continue;
            }

            int instrumentCount = includeInstruments.Count;

            if (instrumentType.Equals("FECs", StringComparison.OrdinalIgnoreCase))
            {
                for (int i = 0; i < instrumentCount; i++)
                {
                    if (forwardPoints is null)
                    {
                        return CommonUtils.DExcelErrorMessage("FEC forward points missing.");
                    }

                    if (tenors is null)
                    {
                        return CommonUtils.DExcelErrorMessage("FEC tenors missing.");
                    }

                    if (fixingDays is null)
                    {
                        return CommonUtils.DExcelErrorMessage("FEC fixing days missing.");
                    }

                    if (includeInstruments[i])
                    {
                        QL.JointCalendar jointCalendar =
                            new(baseCurrencyIndex.fixingCalendar(), quoteCurrencyIndex.fixingCalendar());

                        // In the case of USDZAR, for example, the collateral curve would be the USD Swap curve.
                        rateHelpers.Add(
                            new QL.FxSwapRateHelper(
                                fwdPoint: new QL.QuoteHandle(new QL.SimpleQuote(forwardPoints[i])),
                                tenor: new QL.Period(tenors[i]),
                                fixingDays: (uint) fixingDays[i],
                                calendar: jointCalendar,
                                convention: baseCurrencyIndex.businessDayConvention(),
                                endOfMonth: baseCurrencyIndex.endOfMonth(),
                                spotFx: new QL.QuoteHandle(new QL.SimpleQuote((double) spotFx)),
                                isFxBaseCurrencyCollateralCurrency: true,
                                collateralCurve: baseCurrencyDiscountCurve));
                    }
                }
            }
            else if (instrumentType.Equals("Cross Currency Swaps", StringComparison.OrdinalIgnoreCase))
            {
                for (int i = 0; i < instrumentCount; i++)
                {
                    if (basisSpreads is null)
                    {
                        return CommonUtils.DExcelErrorMessage("Cross currency swap basis spreads missing.");
                    }

                    if (tenors is null)
                    {
                        return CommonUtils.DExcelErrorMessage("Cross currency swap tenors missing.");
                    }

                    if (fixingDays is null)
                    {
                        return CommonUtils.DExcelErrorMessage("Cross currency swap fixing days missing.");
                    }
                    
                    if (includeInstruments[i])
                    {
                        // In the case of USDZAR, for example, the collateral curve would be the USD Swap curve.
                        QL.JointCalendar jointCalendar = 
                            new(baseCurrencyIndex.fixingCalendar(), quoteCurrencyIndex.fixingCalendar());
                        
                        rateHelpers.Add(
                            new QL.ConstNotionalCrossCurrencyBasisSwapRateHelper(
                                basis: new QL.QuoteHandle(new QL.SimpleQuote(basisSpreads[i])),
                                tenor: new QL.Period(tenors[i]),
                                fixingDays: (uint)fixingDays[i],
                                calendar: jointCalendar,
                                convention: baseCurrencyIndex.businessDayConvention(),
                                endOfMonth: false,
                                baseCurrencyIndex: baseCurrencyIndex,
                                quoteCurrencyIndex: quoteCurrencyIndex,
                                collateralCurve: baseCurrencyDiscountCurve,
                                isFxBaseCurrencyCollateralCurrency: true,
                                isBasisOnFxBaseCurrencyLeg: false));
                    }
                }
            }
        }

        QL.YieldTermStructure? termStructure =
            dExcel.InterestRates.CurveBootstrapper.BootstrapCurveFromRateHelpers(
                rateHelpers: rateHelpers, 
                referenceDate: baseDate, 
                dayCountConvention: quoteCurrencyIndex.dayCounter(), 
                interpolation: interpolation);
        
        if (termStructure is null)
        {
            return CommonUtils.DExcelErrorMessage($"Unknown interpolation method: '{interpolation}'");
        }
        
        if ((bool)allowExtrapolation)
        {
            termStructure.enableExtrapolation();
        }
        
        CurveDetails curveDetails = new(termStructure, quoteCurrencyIndex.dayCounter(), interpolation, null,  null, instrumentGroups);
        DataObjectController dataObjectController = DataObjectController.Instance;
        return dataObjectController.Add(handle, curveDetails);
    }
}
