using dExcel.CommonEnums;
using dExcel.Dates;
using dExcel.ExcelUtils;
using dExcel.InterestRates;
using dExcel.Utilities;
using ExcelDna.Integration;
using Omicron;
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
        Name = "d.Curve_BootstrapFxBasisAdjustedCurve",
        Description = "Bootstraps an FX basis adjusted curve.",
        Category = "∂Excel: FX",
        IsVolatile = true,
        IsMacroType = true)]
    public static string BootstrapFxBasisAdjustedCurve(
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
            ExcelTableUtils.GetTableValue<DateTime>(curveParameters, "Value", "Base Date", columnHeaderIndex);
        
        if (baseDate == default)
        {
            return nameof(baseDate).CurveParameterMissingErrorMessage();
        }

        QL.Settings.instance().setEvaluationDate(baseDate.ToQuantLibDate());

        string? baseCurrencyIndexName =
            ExcelTableUtils.GetTableValue<string?>(curveParameters, "Value", "Base Currency Index Name", columnHeaderIndex);
        
        if (baseCurrencyIndexName is null && customBaseCurrencyIndex is null)
        {
            return nameof(baseCurrencyIndexName).CurveParameterMissingErrorMessage();
        }
        
        string? quoteCurrencyIndexName =
            ExcelTableUtils.GetTableValue<string?>(curveParameters, "Value", "Quote Currency Index Name", columnHeaderIndex);
      
        if (quoteCurrencyIndexName is null && customBaseCurrencyIndex is null)
        {
            return nameof(quoteCurrencyIndexName).CurveParameterMissingErrorMessage();
        }

        string? baseCurrencyIndexTenor =
            ExcelTableUtils.GetTableValue<string?>(curveParameters, "Value", "Base Currency Index Tenor", columnHeaderIndex);
        
        if (baseCurrencyIndexTenor is null && 
            customBaseCurrencyIndex is null && 
            quoteCurrencyIndexName != RateIndices.FEDFUND.ToString())
        {
            return nameof(baseCurrencyIndexTenor).CurveParameterMissingErrorMessage();
        }

        string? quoteCurrencyIndexTenor =
            ExcelTableUtils.GetTableValue<string?>(curveParameters, "Value", "Quote Currency Index Tenor", columnHeaderIndex);
        
        if (quoteCurrencyIndexTenor is null && 
            customBaseCurrencyIndex is null && 
            quoteCurrencyIndexName != RateIndices.FEDFUND.ToString())
        {
            return nameof(quoteCurrencyIndexTenor).CurveParameterMissingErrorMessage();
        }
        
        string? interpolation =
            ExcelTableUtils.GetTableValue<string?>(curveParameters, "Value", "Interpolation", columnHeaderIndex);
        
        if (interpolation is null)
        {
            return nameof(interpolation).CurveParameterMissingErrorMessage();
        }
        
        double? spotFx =
            ExcelTableUtils.GetTableValue<double?>(curveParameters, "Value", "Spot FX", columnHeaderIndex);
        
        if (spotFx is null)
        {
            return nameof(spotFx).CurveParameterMissingErrorMessage();
        }
        
        bool allowExtrapolation =
            ExcelTableUtils.GetTableValue<bool?>(curveParameters, "Value", "Allow Extrapolation", columnHeaderIndex) ??
            false;

        string? baseCurrencyDiscountCurve =
            ExcelTableUtils.GetTableValue<string?>(
                table: curveParameters, 
                columnHeader: "Value", 
                rowHeader: "Base Currency Discount Curve", 
                rowIndexOfColumnHeaders: columnHeaderIndex);

        if (baseCurrencyDiscountCurve is null)
        {
            return nameof(baseCurrencyDiscountCurve).CurveParameterMissingErrorMessage(); 
        }
        
        QL.RelinkableYieldTermStructureHandle baseCurrencyDiscountCurveTermStructure = new();
        QL.YieldTermStructure? baseCurrencyYieldTermStructure = 
            CurveUtils.GetCurveObject(baseCurrencyDiscountCurve);
        
        baseCurrencyDiscountCurveTermStructure.linkTo(baseCurrencyYieldTermStructure);
        
        (QL.RelinkableYieldTermStructureHandle? baseCurrencyForecastCurve,
                string? baseCurrencyForecastCurveErrorMessage) =
           CurveUtils.GetYieldTermStructure(
               yieldTermStructureName: "BaseCurrencyForecastCurve", 
               table: curveParameters, 
               columnHeaderIndex: columnHeaderIndex, 
               allowExtrapolation: allowExtrapolation);

        if (baseCurrencyForecastCurveErrorMessage is not null)
        {
            return baseCurrencyForecastCurveErrorMessage;
        }
        
        (QL.RelinkableYieldTermStructureHandle? quoteCurrencyForecastCurve,
                string? quoteCurrencyForecastCurveErrorMessage) =
           CurveUtils.GetYieldTermStructure(
               yieldTermStructureName: "QuoteCurrencyForecastCurve", 
               table: curveParameters, 
               columnHeaderIndex: columnHeaderIndex, 
               allowExtrapolation: allowExtrapolation);

        if (quoteCurrencyForecastCurveErrorMessage is not null)
        {
            return quoteCurrencyForecastCurveErrorMessage;
        }
      
        QL.IborIndex? quoteCurrencyIndex = (customQuoteCurrencyIndex is null) ? 
            dExcel.InterestRates.CurveBootstrapper.GetIborIndex(
                indexName: quoteCurrencyIndexName, 
                indexTenor: quoteCurrencyIndexTenor, 
                forecastCurve: quoteCurrencyForecastCurve) : 
            null;

        if (quoteCurrencyIndex is null)
        {
            string? indexName = ExcelTableUtils.GetTableValue<string>(customQuoteCurrencyIndex, "Value", "Name");
            string? tenor = ExcelTableUtils.GetTableValue<string>(customQuoteCurrencyIndex, "Value", "Tenor");
            int? customSettlementDays = ExcelTableUtils.GetTableValue<int>(customQuoteCurrencyIndex, "Value", "Settlement Days");
            string? customCurrencyToParse = ExcelTableUtils.GetTableValue<string>(customQuoteCurrencyIndex, "Value", "Currency");
            ParserUtils.TryParseQuantLibCurrency(
                customCurrencyToParse, 
                out QL.Currency? customCurrency,
                out string? customCurrencyErrorMessage);

            string? calendarsToParse = ExcelTableUtils.GetTableValue<string>(customQuoteCurrencyIndex, "Value", "Calendars");
            DateUtils.TryParseCalendars(
                calendarsToParse, 
                out QL.Calendar? customCalendar,
                out string customCalendarErrorMessage);

            string? businessDayConventionToParse =
                ExcelTableUtils.GetTableValue<string>(customQuoteCurrencyIndex, "Value", "Business Day Convention");
            (QL.BusinessDayConvention? customBusinessDayConvention, string? customBusinessDayConventionErrorMessage) =
                DateUtils.ParseBusinessDayConvention(businessDayConventionToParse);

            string? dayCountConventionToParse =
                ExcelTableUtils.GetTableValue<string>(customQuoteCurrencyIndex, "Value", "Day Count Convention");

            QL.DayCounter? customDayCountConvention = DateUtils.ParseDayCountConvention(dayCountConventionToParse);
            quoteCurrencyIndex = new QL.IborIndex(indexName, new QL.Period(tenor), (int)customSettlementDays, customCurrency,
                customCalendar, (QL.BusinessDayConvention)customBusinessDayConvention, false, customDayCountConvention);

        }

        QL.IborIndex? baseCurrencyIndex = (customBaseCurrencyIndex is null) ?
            dExcel.InterestRates.CurveBootstrapper.GetIborIndex(
                indexName: baseCurrencyIndexName, 
                indexTenor: baseCurrencyIndexTenor, 
                forecastCurve: baseCurrencyForecastCurve) :
            null;

        if (baseCurrencyIndex is null)
        {
            string? indexName = ExcelTableUtils.GetTableValue<string>(customBaseCurrencyIndex, "Value", "Name");
            string? tenor = ExcelTableUtils.GetTableValue<string>(customBaseCurrencyIndex, "Value", "Tenor");
            int? customSettlementDays = ExcelTableUtils.GetTableValue<int>(customBaseCurrencyIndex, "Value", "Settlement Days");
            string? customCurrencyToParse = ExcelTableUtils.GetTableValue<string>(customBaseCurrencyIndex, "Value", "Currency");
            ParserUtils.TryParseQuantLibCurrency(
                customCurrencyToParse, 
                out QL.Currency? customCurrency,
                out string? customCurrencyErrorMessage);

            string? calendarsToParse = ExcelTableUtils.GetTableValue<string>(customBaseCurrencyIndex, "Value", "Calendars");
            DateUtils.TryParseCalendars(
                calendarsToParse, 
                out QL.Calendar? customCalendar,
                out string customCalendarErrorMessage);

            string? businessDayConventionToParse =
                ExcelTableUtils.GetTableValue<string>(customBaseCurrencyIndex, "Value", "Business Day Convention");
            (QL.BusinessDayConvention? customBusinessDayConvention, string? customBusinessDayConventionErrorMessage) =
                DateUtils.ParseBusinessDayConvention(businessDayConventionToParse);

            string? dayCountConventionToParse =
                ExcelTableUtils.GetTableValue<string>(customBaseCurrencyIndex, "Value", "Day Count Convention");

            QL.DayCounter? customDayCountConvention = DateUtils.ParseDayCountConvention(dayCountConventionToParse);
            baseCurrencyIndex = new QL.IborIndex(indexName, new QL.Period(tenor), (int)customSettlementDays, customCurrency,
                customCalendar, (QL.BusinessDayConvention)customBusinessDayConvention, false, customDayCountConvention);
        }

        if (baseCurrencyIndex is null)
        {
            return CommonUtils.DExcelErrorMessage($"Unsupported rate index: {baseCurrencyIndex}");
        }

        if (quoteCurrencyIndex is null)
        {
            return CommonUtils.DExcelErrorMessage($"Unsupported rate index: {quoteCurrencyIndexName}");
        }
        
        QL.RateHelperVector rateHelpers = new();
        bool instrumentsWithNaNsFound = false;
        
        foreach (object instrumentGroup in instrumentGroups)
        {
            object[,] instruments = (object[,]) instrumentGroup;
            string? instrumentType = ExcelTableUtils.GetTableLabel(instruments);
            if (instrumentType is null)
            {
                return CommonUtils.DExcelErrorMessage("No instrument type found.");
            }
            
            List<string>? tenors = ExcelTableUtils.GetColumn<string>(instruments, "Tenors");
            List<double>? basisSpreads = ExcelTableUtils.GetColumn<double>(instruments, "Basis Spreads");
            List<double>? forwardPoints = ExcelTableUtils.GetColumn<double>(instruments, "Forward Points");
            List<int>? fixingDays = ExcelTableUtils.GetColumn<int>(instruments, "Fixing Days");
            List<bool>? includeInstruments = ExcelTableUtils.GetColumn<bool>(instruments, "Include");

            if (includeInstruments is null)
            {
                continue;
            }

            int instrumentCount = includeInstruments.Count;

            if (instrumentType.IgnoreCaseEquals("FECs", "FX Forwards"))
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

                    instrumentsWithNaNsFound = instrumentsWithNaNsFound || double.IsNaN(forwardPoints[i]);
                    
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
                                collateralCurve: baseCurrencyDiscountCurveTermStructure));
                    }
                }
            }
            else if (instrumentType.IgnoreCaseEquals("Cross Currency Swap", "Cross Currency Swaps"))
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
                    
                    instrumentsWithNaNsFound = instrumentsWithNaNsFound || double.IsNaN(basisSpreads[i]);
                    
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
                                collateralCurve: baseCurrencyDiscountCurveTermStructure,
                                isFxBaseCurrencyCollateralCurrency: true,
                                isBasisOnFxBaseCurrencyLeg: false));
                    }
                }
            }
            else
            {
                return CommonUtils.DExcelErrorMessage($"Invalid instrument type: '{instrumentType}'");
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
        
        if (allowExtrapolation)
        {
            termStructure.enableExtrapolation();
        }
        
        CurveDetails curveDetails = 
            new(termStructure, quoteCurrencyIndex.dayCounter(), interpolation, null,  null, instrumentGroups);
        
        DataObjectController dataObjectController = DataObjectController.Instance;
        string warningMessage = instrumentsWithNaNsFound ? "Instruments with NaNs found" : "";
        return dataObjectController.Add(handle, curveDetails, warningMessage);
    }
    
    /// <summary>
    /// Extracts and bootstraps an FX basis adjusted curve from the Omicron database.
    /// </summary>
    /// <param name="handle">The 'handle' or name used to refer to the object in memory.
    /// Each object in a workbook must have a a unique handle.</param>
    /// <param name="curveName">The name of the curve in Omicron. Current available options are:
    /// 'ZAR-Swap', 'USD-OIS'</param>
    /// <param name="baseDate"></param>
    /// <param name="interpolation"></param>
    /// <returns></returns>
    [ExcelFunction(
        Name = "d.Curve_GetFxBasisAdjustedCurve",
        Description = "Extracts and bootstraps an FX basis adjusted curve from the Omicron database.",
        Category = "∂Excel: Interest Rates")]
    public static string GetFxBasisAdjustedCurve(
        [ExcelArgument(Name = "Handle", Description = DescriptionUtils.Handle)]
        string handle,
        [ExcelArgument(
            Name = "Curve Name",
            Description = 
                "The name of the curve in Omicron. Current available options are.")]
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
        
        // TODO: List the available types of interpolation.
        // One could create more complicated abstract code for mapping from quotes to 2d tables but I would advise
        // against this.
        string baseIndexName;
        Tenor baseIndexTenor;
        Currency baseCurrency;
        string spreadIndexName;
        Tenor spreadIndexTenor;
        Currency spreadCurrency;
        string quoteCurrencyForecastCurveHandle;
        string baseCurrencyForecastCurveHandle;
        string baseCurrencyDiscountCurveHandle;
        
        if (curveName.IgnoreCaseEquals(OmicronFxBasisCurves.USDZAR.ToString()))
        {
               baseIndexName = RateIndices.JIBAR.ToString();
               baseIndexTenor = new Tenor(3, TenorUnit.Month);
               baseCurrency = Currency.ZAR;
               spreadIndexName = RateIndices.USD_LIBOR.ToString();
               spreadIndexTenor = new Tenor(3, TenorUnit.Month);
               spreadCurrency = Currency.USD;
               quoteCurrencyForecastCurveHandle = 
                   InterestRates.CurveBootstrapper.Get(handle + "_ZAR", "ZAR-Swap", baseDate);
               baseCurrencyDiscountCurveHandle = 
                   InterestRates.CurveBootstrapper.Get(handle + "_USDOIS", OmicronSwapCurves.USD_OIS.ToString(), baseDate);
               baseCurrencyForecastCurveHandle = 
                   InterestRates.CurveBootstrapper.Get(handle + "_USDSwap", OmicronSwapCurves.USD_Swap.ToString(), baseDate);
               
        }
        else if (curveName.IgnoreCaseEquals(OmicronFxBasisCurves.USDZAR_SOFR.ToString()))
        {
               baseIndexName = RateIndices.JIBAR.ToString();
               baseIndexTenor = new Tenor(3, TenorUnit.Month);
               baseCurrency = Currency.ZAR;
               spreadIndexName = RateIndices.SOFR.ToString();
               spreadIndexTenor = new Tenor(3, TenorUnit.Month);
               spreadCurrency = Currency.USD;
               quoteCurrencyForecastCurveHandle = 
                   InterestRates.CurveBootstrapper.Get(handle + "_ZAR", "ZAR-Swap", baseDate);
               baseCurrencyDiscountCurveHandle = 
                   InterestRates.CurveBootstrapper.Get(handle + "_SOFR", OmicronSwapCurves.SOFR.ToString(), baseDate);
               baseCurrencyForecastCurveHandle = baseCurrencyDiscountCurveHandle;
        }
        else
        {
            return CommonUtils.DExcelErrorMessage($"Unsupported curve name: {curveName}");
        }

        List<QuoteValue> quoteValues;
        try
        {
            quoteValues =
                Task.Run(
                    function: () => 
                        OmicronUtils.OmicronUtils.GetAllFxBasisCurveQuotes(
                            spreadIndexName: spreadIndexName,
                            spreadIndexTenor: spreadIndexTenor,
                            baseIndexName: baseIndexName,
                            baseIndexTenor: baseIndexTenor,
                            numeratorCurrency: baseCurrency,
                            denominatorCurrency: spreadCurrency,
                            marketDataDate: baseDate)).Result;
        }
        catch (Exception ex)
        {
            if (!NetworkUtils.GetVpnConnectionStatus())
            {
                return CommonUtils.DExcelErrorMessage("Not connected to Deloitte network/VPN.");
            }

            return CommonUtils.DExcelErrorMessage($"Unknown error. {ex.Message}");
        }

        double fxSpot = (quoteValues.Where(x => x.Type == new FxSpot(baseCurrency, spreadCurrency)).ElementAt(0)).Value;
        object[,] curveParameters =
        {
            {"CurveUtils Parameters", ""},
            {"Parameter", "Value"},
            {"BaseDate", baseDate.ToOADate()},
            {"BaseCurrencyIndexName", spreadIndexName},
            {"BaseCurrencyIndexTenor", spreadIndexTenor.ToString()}, 
            {"BaseCurrencyDiscountCurve", baseCurrencyDiscountCurveHandle},
            {"BaseCurrencyForecastCurve", baseCurrencyForecastCurveHandle},  
            {"QuoteCurrencyIndexName", baseIndexName},
            {"QuoteCurrencyIndexTenor", baseIndexTenor.ToString()},
            {"QuoteCurrencyForecastCurve", quoteCurrencyForecastCurveHandle},  
            {"SpotFx", fxSpot},
            {"Interpolation", interpolation},
        };

        
        List<QuoteValue> fxForwards = quoteValues.Where(x => x.Type.GetType() == typeof(FxForward)).ToList();
        fxForwards = fxForwards.OrderBy(x => ((FxForward)x.Type).Tenor, new TenorComparer()).ToList();
        object[,] fxForwardInstruments = new object[fxForwards.Count + 2, 4];
        fxForwardInstruments[0, 0] = "FX Forwards";
        fxForwardInstruments[1, 0] = "Tenors";
        fxForwardInstruments[1, 1] = "ForwardPoints";
        fxForwardInstruments[1, 2] = "FixingDays";
        fxForwardInstruments[1, 3] = "Include";

        int fxForwardRow = 2;
        foreach (QuoteValue fxForward in fxForwards)
        {
            fxForwardInstruments[fxForwardRow, 0] = ((FxForward) fxForward.Type).Tenor.ToString();
            fxForwardInstruments[fxForwardRow, 1] = fxForward.Value;
            fxForwardInstruments[fxForwardRow, 2] = 2; // TODO: Map this somewhere.
            fxForwardInstruments[fxForwardRow, 3] = "TRUE";
            fxForwardRow++;
        }
        
        List<QuoteValue> crossCurrencySwaps = quoteValues.Where(x => x.Type.GetType() == typeof(FxBasisSwap)).ToList();
        crossCurrencySwaps = crossCurrencySwaps.OrderBy(x => ((FxBasisSwap)x.Type).Tenor, new TenorComparer()).ToList();
        object[,] crossCurrencySwapInstruments = new object[crossCurrencySwaps.Count + 2, 4];
        crossCurrencySwapInstruments[0, 0] = "Cross Currency Swaps";
        crossCurrencySwapInstruments[1, 0] = "Tenors";
        crossCurrencySwapInstruments[1, 1] = "BasisSpreads";
        crossCurrencySwapInstruments[1, 2] = "FixingDays";
        crossCurrencySwapInstruments[1, 3] = "Include";

        int swapRow = 2;
        foreach (QuoteValue crossCurrencySwap in crossCurrencySwaps)
        {
            crossCurrencySwapInstruments[swapRow, 0] = ((FxBasisSwap) crossCurrencySwap.Type).Tenor.ToString();
            crossCurrencySwapInstruments[swapRow, 1] = crossCurrencySwap.Value;
            crossCurrencySwapInstruments[swapRow, 2] = 2; // TODO: Map this somewhere.
            crossCurrencySwapInstruments[swapRow, 3] = "TRUE";
            swapRow++;
        }
        
        List<object> instruments = new();
        if (fxForwards.Any()) instruments.Add(fxForwardInstruments);
        if (crossCurrencySwaps.Any()) instruments.Add(crossCurrencySwapInstruments);

        return BootstrapFxBasisAdjustedCurve(handle, curveParameters, null, null, instruments.ToArray());
    }
}
