namespace dExcel.Curves;

using ExcelDna.Integration;
using ExcelUtils;
using QLNet;
using Utilities;

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
        DateTime baseDate = ExcelTable.GetTableValue<DateTime>(curveParameters, "Value", "BaseDate", 1);
        if (baseDate == default)
        {
            return $"{CommonUtils.DExcelErrorPrefix} Base date missing from curve parameters.";
        }
        
        Settings.setEvaluationDate(baseDate);

        string? rateIndexName = ExcelTable.GetTableValue<string>(curveParameters, "Value", "RateIndexName");
        if (rateIndexName is null && customRateIndex is null)
        {
            return $"{CommonUtils.DExcelErrorPrefix} Rate index missing from curve parameters (and no custom rate index provided).";
        }

        string? indexTenor = ExcelTable.GetTableValue<string>(curveParameters, "Value", "RateIndexTenor");
        if (indexTenor is null && customRateIndex is null && rateIndexName != "FEDFUND")
        {
            return $"{CommonUtils.DExcelErrorPrefix} Please provide a rate index tenor in the curve parameters.";
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
            return $"{CommonUtils.DExcelErrorPrefix} Unsupported rate index: {rateIndexName}";
        }

        List<RateHelper> rateHelpers = new();

        foreach (object instrumentGroup in instrumentGroups)
        {
            object[,] instruments = (object[,])instrumentGroup;

            // TODO: Make this case insensitive.
            string? instrumentType = ExcelTable.GetTableLabel(instruments);
            List<string>? tenors = ExcelTable.GetColumn<string>(instruments, "Tenors");
            List<string>? fraTenors = ExcelTable.GetColumn<string>(instruments, "FraTenors");
            List<DateTime>? endDates = ExcelTable.GetColumn<DateTime>(instruments, "EndDates");
            List<double>? rates = ExcelTable.GetColumn<double>(instruments, "Rates");
            List<bool>? includeInstruments = ExcelTable.GetColumn<bool>(instruments, "Include");

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
                        return $"{CommonUtils.DExcelErrorPrefix} Deposit rates missing.";
                    }

                    if (tenors is null)
                    {
                        return $"{CommonUtils.DExcelErrorPrefix} Deposit tenors missing";
                    }
                    
                    if (includeInstruments[i])
                    {
                        rateHelpers.Add(
                            item: new DepositRateHelper(
                            rate: rates[i],
                            tenor: new Period(tenors?[i]),
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
                            return $"{CommonUtils.DExcelErrorPrefix} FRA tenors missing.";
                        }

                        if (rates is null)
                        {
                            return $"{CommonUtils.DExcelErrorPrefix} FRA rates missing.";
                        }
                        
                        rateHelpers.Add(
                            item: new FraRateHelper(
                            rate: new Handle<Quote>(new SimpleQuote(rates?[i])),
                            periodToStart: new Period(fraTenors?[i]),
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
                        return $"{CommonUtils.DExcelErrorPrefix} Swap rates missing.";
                    }

                    if (tenors is null)
                    {
                        return $"{CommonUtils.DExcelErrorPrefix} Swap tenors missing";
                    }
                    
                    if (includeInstruments[i])
                    {
                        rateHelpers.Add(
                            item: new SwapRateHelper(
                            rate: new Handle<Quote>(new SimpleQuote(rates?[i])),
                            tenor: new Period(tenors?[i]),
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

        YieldTermStructure termStructure =
            new PiecewiseYieldCurve<Discount, LogLinear>(
               referenceDate: new Date(baseDate),
               instruments: rateHelpers,
               dayCounter: rateIndex.dayCounter(),
               jumps: new List<Handle<Quote>>(),
               jumpDates: new List<Date>(),
               accuracy: 1.0e-20);
        
        Dictionary<string, object> curveDetails = new()
        {
            // TODO: Add instruments used in bootstrapping.
            // TODO: Get out anchor dates.
            ["CurveUtils.Object"] = termStructure,
        };
        
        return DataObjectController.Add(handle, curveDetails);
    }
}
