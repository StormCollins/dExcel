using dExcel.Dates;

namespace dExcel.InterestRates;

using ExcelUtils;
using ExcelDna.Integration;
using QLNet;
using Utilities;

/// <summary>
/// A class for bootstrapping single curves e.g., the ZAR swap curve or USD OIS swap curve.
/// </summary>
public static class SingleCurveBootstrapper
{
    /// <summary>
    /// Bootstraps a single curve i.e., this is not a multi-curve bootstrapper.
    /// Available Indices: EURIBOR, FEDFUND (OIS), JIBAR, USD-LIBOR.
    /// </summary>
    /// <param name="handle"></param>
    /// <param name="curveParameters"></param>
    /// <param name="customRateIndex"></param>
    /// <param name="instrumentGroups"></param>
    /// <returns></returns>
    [ExcelFunction(
        Name = "d.Curve_SingleCurveBootstrap",
        Description = "Bootstraps a single curve i.e., this is not a multi-curve bootstrapper.\n" +
                      "Available Indices: EURIBOR, FEDFUND (OIS), JIBAR, USD-LIBOR",
        Category = "∂Excel: Interest Rates")]
    public static string Bootstrap(
        [ExcelArgument(
            Name = "Handle", 
            Description = 
                "The 'handle' or name used to refer to the object in memory.\n" + 
                "Each curve must have a a unique handle.")]
        string handle, 
        [ExcelArgument(
            Name = "Curve Parameters", 
            Description = "The curves parameters: 'BaseDate', 'RateIndexName', 'RateIndexTenor', 'Interpolation'.")]
        object[,] curveParameters, 
        [ExcelArgument(
            Name = "(Optional)Custom Rate Index",
            Description = 
                "Only populate this parameter if you have not supplied a 'RateIndexName' in the curve parameters.")]
        object[,]? customRateIndex = null, 
        [ExcelArgument(
            Name = "Instrument Groups",
            Description = "The instrument groups used to bootstrap the curve e.g., 'Deposits', 'FRAs', 'Swaps'.")]
        params object[] instrumentGroups)
    {
        int columnHeaderIndex = ExcelTableUtils.GetRowIndex(curveParameters, "Parameter");
        DateTime baseDate = ExcelTableUtils.GetTableValue<DateTime>(curveParameters, "Value", "BaseDate", columnHeaderIndex);
        if (baseDate == default)
        {
            return CommonUtils.CurveParameterMissingErrorMessage(nameof(baseDate).ToUpper());
        }
        
        Settings.setEvaluationDate(baseDate);

        string? rateIndexName = ExcelTableUtils.GetTableValue<string>(curveParameters, "Value", "RateIndexName", columnHeaderIndex);
        if (rateIndexName is null && customRateIndex is null)
        {
            return CommonUtils.CurveParameterMissingErrorMessage(nameof(rateIndexName).ToUpper());
        }

        string? rateIndexTenor = ExcelTableUtils.GetTableValue<string>(curveParameters, "Value", "RateIndexTenor", columnHeaderIndex);
        if (rateIndexTenor is null && customRateIndex is null && rateIndexName != "FEDFUND")
        {
            return CommonUtils.CurveParameterMissingErrorMessage(nameof(rateIndexTenor).ToUpper());
        }

        string? interpolation = ExcelTableUtils.GetTableValue<string>(curveParameters, "Value", "Interpolation", columnHeaderIndex);
        if (interpolation == null)
        {
            return CommonUtils.CurveParameterMissingErrorMessage(nameof(rateIndexTenor).ToUpper());
        }

        IborIndex? rateIndex = null;
        if (rateIndexName is not null)
        {
            rateIndex =
                rateIndexName switch
                {
                    "EURIBOR" => new Euribor(new Period(rateIndexTenor)),
                    "FEDFUND" => new FedFunds(),
                    "JIBAR" => new Jibar(new Period(rateIndexTenor)),
                    "USD-LIBOR" => new USDLibor(new Period(rateIndexTenor)),
                    _ => null,
                };
        }
        // else
        // {
        //     string? tenor = ExcelTable.GetTableValue<string>(customRateIndex, "Value", "Tenor");
        //     int? settlementDays = ExcelTable.GetTableValue<int>(customRateIndex, "Value", "SettlementDay");
        //     string? currency = ExcelTable.GetTableValue<string>(customRateIndex, "Value", "Currency");
        //     (Currency x, Calendar y) z =
        //         currency switch
        //         {
        //             // TODO: Use reflection here.
        //             "USD" => (new USDCurrency(), new UnitedStates()),
        //             "ZAR" => (new ZARCurrency(), new SouthAfrica()),
        //             _ => throw new NotImplementedException(),
        //         };
        //
        //     
        //
        //     rateIndex = new IborIndex("Test", new Period(tenor), settlementDays, z.x, z.y, )
        // }


        if (rateIndex is null)
        {
            return CommonUtils.DExcelErrorMessage($"Unsupported rate index: {rateIndexName}");
        }

        List<RateHelper> rateHelpers = new();

        foreach (object instrumentGroup in instrumentGroups)
        {
            object[,] instruments = (object[,])instrumentGroup;
            string? instrumentType = ExcelTableUtils.GetTableLabel(instruments);
            List<string>? tenors = ExcelTableUtils.GetColumn<string>(instruments, "Tenors");
            List<string>? fraTenors = ExcelTableUtils.GetColumn<string>(instruments, "FraTenors");
            List<DateTime>? endDates = ExcelTableUtils.GetColumn<DateTime>(instruments, "EndDates");
            List<double>? rates = ExcelTableUtils.GetColumn<double>(instruments, "Rates");
            List<bool>? includeInstruments = ExcelTableUtils.GetColumn<bool>(instruments, "Include");

            if (includeInstruments is null)
            {
                continue;
            }
                
            int instrumentCount = includeInstruments.Count;

            if (string.Equals(instrumentType, "Deposits", StringComparison.OrdinalIgnoreCase))
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
                    
                    if (includeInstruments[i])
                    {
                        rateHelpers.Add(
                            item: new DepositRateHelper(
                            rate: rates[i],
                            tenor: new Period(tenors[i]),
                            fixingDays: rateIndex.fixingDays(),
                            calendar: rateIndex.fixingCalendar(),
                            convention: rateIndex.businessDayConvention(),
                            endOfMonth: rateIndex.endOfMonth(),
                            dayCounter: rateIndex.dayCounter()));
                    }
                }
            }
            else if (string.Compare(instrumentType, "FRAs", StringComparison.InvariantCultureIgnoreCase) == 0)
            {
                for (int i = 0; i < instrumentCount; i++)
                {
                    if (includeInstruments[i])
                    {
                        if (fraTenors is null)
                        {
                            return CommonUtils.DExcelErrorMessage("FRA tenors missing.");
                        }

                        if (rates is null)
                        {
                            return CommonUtils.DExcelErrorMessage("FRA rates missing.");
                        }
                        
                        rateHelpers.Add(
                            item: new FraRateHelper(
                            rate: new Handle<Quote>(new SimpleQuote(rates[i])),
                            periodToStart: new Period(fraTenors[i]),
                            iborIndex: rateIndex));
                    }
                }
            }
            else if (string.Compare(instrumentType, "Interest Rate Swaps", StringComparison.InvariantCultureIgnoreCase) == 0)
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
                    
                    if (includeInstruments[i])
                    {
                        rateHelpers.Add(
                            item: new SwapRateHelper(
                            rate: new Handle<Quote>(new SimpleQuote(rates[i])),
                            tenor: new Period(tenors[i]),
                            calendar: rateIndex.fixingCalendar(),
                            fixedFrequency: Frequency.Quarterly,
                            fixedConvention: rateIndex.businessDayConvention(),
                            fixedDayCount: rateIndex.dayCounter(),
                            iborIndex: rateIndex));
                    }
                }
            }
            else if (string.Compare(instrumentType, "OISs", StringComparison.InvariantCultureIgnoreCase) == 0)
            {
                for (int i = 0; i < instrumentCount; i++)
                {
                    if (includeInstruments[i])
                    {
                        // rateHelpers.Add(
                        //     item: new DatedOISRateHelper(
                        //     startDate: (DateTime)DateUtils.AddTenorToDate(baseDate, "2d", "USD", "ModFol"), 
                        //     endDate: endDates?[i],
                        //     fixedRate: new Handle<Quote>(new SimpleQuote(rates?[i])),
                        //     overnightIndex: rateIndex as OvernightIndex));

                        rateHelpers.Add(
                            item: new OISRateHelper(2, new Period(tenors?[i]), new Handle<Quote>(new SimpleQuote(rates?[i])), rateIndex as OvernightIndex));
                    }
                }
            }
        }

        YieldTermStructure termStructure;
        if (string.Compare(interpolation, "BackwardFlat", StringComparison.InvariantCultureIgnoreCase) == 0)
        {
            termStructure =
                new PiecewiseYieldCurve<Discount, BackwardFlat>(
                   referenceDate: new Date(baseDate),
                   instruments: rateHelpers,
                   dayCounter: rateIndex.dayCounter(),
                   jumps: new List<Handle<Quote>>(),
                   jumpDates: new List<Date>(),
                   accuracy: 1.0e-20);
        }
        else if (string.Compare(interpolation, "Cubic", StringComparison.InvariantCultureIgnoreCase) == 0)
        {
            termStructure =
                new PiecewiseYieldCurve<Discount, Cubic>(
                   referenceDate: new Date(baseDate),
                   instruments: rateHelpers,
                   dayCounter: rateIndex.dayCounter(),
                   jumps: new List<Handle<Quote>>(),
                   jumpDates: new List<Date>(),
                   accuracy: 1.0e-20);
        }
        else if (string.Compare(interpolation, "Exponential", StringComparison.InvariantCultureIgnoreCase) == 0)
        {
            termStructure =
                new PiecewiseYieldCurve<Discount, LogLinear>(
                   referenceDate: new Date(baseDate),
                   instruments: rateHelpers,
                   dayCounter: rateIndex.dayCounter(),
                   jumps: new List<Handle<Quote>>(),
                   jumpDates: new List<Date>(),
                   accuracy: 1.0e-20);
        }
        else if (string.Compare(interpolation, "ForwardFlat", StringComparison.InvariantCultureIgnoreCase) == 0)
        {
            termStructure =
                new PiecewiseYieldCurve<Discount, ForwardFlat>(
                   referenceDate: new Date(baseDate),
                   instruments: rateHelpers,
                   dayCounter: rateIndex.dayCounter(),
                   jumps: new List<Handle<Quote>>(),
                   jumpDates: new List<Date>(),
                   accuracy: 1.0e-20);
        }
        else if (string.Compare(interpolation, "Linear", StringComparison.InvariantCultureIgnoreCase) == 0)
        {
            termStructure =
                new PiecewiseYieldCurve<Discount, Linear>(
                   referenceDate: new Date(baseDate),
                   instruments: rateHelpers,
                   dayCounter: rateIndex.dayCounter(),
                   jumps: new List<Handle<Quote>>(),
                   jumpDates: new List<Date>(),
                   accuracy: 1.0e-20);
        }
        else if (string.Compare(interpolation, "LogCubic", StringComparison.InvariantCultureIgnoreCase) == 0)
        {
            termStructure =
                new PiecewiseYieldCurve<Discount, LogCubic>(
                   referenceDate: new Date(baseDate),
                   instruments: rateHelpers,
                   dayCounter: rateIndex.dayCounter(),
                   jumps: new List<Handle<Quote>>(),
                   jumpDates: new List<Date>(),
                   accuracy: 1.0e-20);
        }
        else
        {
            return CommonUtils.DExcelErrorMessage($"Unknown interpolation method: '{interpolation}'");  
        }
         
        CurveDetails curveDetails = new(termStructure, rateIndex.dayCounter(), interpolation, new List<Date>(), new List<double>());
        DataObjectController dataObjectController = DataObjectController.Instance;
        return dataObjectController.Add(handle, curveDetails);
    }
}
