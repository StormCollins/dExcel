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
    [ExcelFunction(
        Name = "d.Curve_SingleCurveBootstrap",
        Description = "Bootstraps a single curve i.e., this is not a multi-curve bootstrapper.\n" +
                      "Available Indices: EURIBOR, FEDFUND (OIS), JIBAR, USD-LIBOR",
        Category = "∂Excel: Interest Rates")]
    public static string Bootstrap(
        string handle, 
        object[,] curveParameters, 
        object[,]? customRateIndex = null, 
        params object[] instrumentGroups)
    {
        DateTime baseDate = ExcelTableUtils.GetTableValue<DateTime>(curveParameters, "Value", "BaseDate", 1);
        if (baseDate == default)
        {
            return CommonUtils.DExcelErrorMessage($"Curve parameter missing: '{nameof(baseDate).ToUpper()}'.");
        }
        
        Settings.setEvaluationDate(baseDate);

        string? rateIndexName = ExcelTableUtils.GetTableValue<string>(curveParameters, "Value", "RateIndexName", 0);
        if (rateIndexName is null && customRateIndex is null)
        {
            return CommonUtils.DExcelErrorMessage("Rate index missing from curve parameters (and no custom rate index provided).");
        }

        string? indexTenor = ExcelTableUtils.GetTableValue<string>(curveParameters, "Value", "RateIndexTenor", 0);
        if (indexTenor is null && customRateIndex is null && rateIndexName != "FEDFUND")
        {
            return CommonUtils.DExcelErrorMessage("Please provide a rate index tenor in the curve parameters.");
        }

        string? interpolationParameter = ExcelTableUtils.GetTableValue<string>(curveParameters, "Value", "Interpolation", 0);
        if (interpolationParameter == null)
        {
            return CommonUtils.DExcelErrorMessage("'Interpolation' not set in parameters.");
        }

        IborIndex? rateIndex = null;
        if (rateIndexName is not null)
        {
            rateIndex =
                rateIndexName switch
                {
                    "EURIBOR" => new Euribor(new Period(indexTenor)),
                    "FEDFUND" => new FedFunds(),
                    "JIBAR" => new Jibar(new Period(indexTenor)),
                    "USD-LIBOR" => new USDLibor(new Period(indexTenor)),
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
            return CommonUtils.DExcelErrorMessage("Unsupported rate index: {rateIndexName}");
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
                        rateHelpers.Add(
                            item: new DatedOISRateHelper(
                            startDate: baseDate, 
                            endDate: endDates?[i],
                            fixedRate: new Handle<Quote>(new SimpleQuote(rates?[i])),
                            overnightIndex: rateIndex as OvernightIndex));
                    }
                }
            }
        }

        if (!CommonUtils.TryParseInterpolation(
                interpolationMethodToParse: interpolationParameter,
                interpolation: out IInterpolationFactory? interpolation,
                errorMessage: out string? interpolationErrorMessage))
        {
            return interpolationErrorMessage;
        }
        
        Type interpolationType = typeof(PiecewiseYieldCurve).MakeGenericType(typeof(Discount), interpolation.GetType());
        object? termStructure = 
            Activator.CreateInstance(
                interpolationType, 
                new Date(baseDate), 
                rateHelpers, 
                rateIndex.dayCounter(), 
                new List<Handle<Quote>>(), 
                new List<Date>(), 
                1.0e-20);
        
        // YieldTermStructure termStructure =
        //     new PiecewiseYieldCurve<Discount, LogLinear>(
        //        referenceDate: new Date(baseDate),
        //        instruments: rateHelpers,
        //        dayCounter: rateIndex.dayCounter(),
        //        jumps: new List<Handle<Quote>>(),
        //        jumpDates: new List<Date>(),
        //        accuracy: 1.0e-20);
         
        CurveDetails curveDetails = new(termStructure, rateIndex.dayCounter(), interpolation, new List<Date>(), new List<double>());
        return DataObjectController.Add(handle, curveDetails);
    }
}
