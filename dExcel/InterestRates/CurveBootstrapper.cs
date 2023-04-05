namespace dExcel.Curves;

using ExcelUtils;
using ExcelDna.Integration;
using InterestRates;
using Omicron;
using OmicronUtils;
using QLNet;
using Utilities;

public class TenorComparer : Comparer<Tenor>
{
    public override int Compare(Tenor x, Tenor y)
    {
        if (x.Unit == y.Unit)
        {
            return x.Amount.CompareTo(y.Amount);
        }

        int xAmount = x.Unit switch
        {
            TenorUnit.Day => x.Amount,
            TenorUnit.Week => x.Amount * 7,
            TenorUnit.Month => x.Amount * 30,
            TenorUnit.Year => x.Amount * 365,
            _ => x.Amount,
        };

        int yAmount = y.Unit switch
        {
            TenorUnit.Day => y.Amount,
            TenorUnit.Week => y.Amount * 7,
            TenorUnit.Month => y.Amount * 30,
            TenorUnit.Year => y.Amount * 365,
            _ => y.Amount,
        };
        
        return xAmount.CompareTo(yAmount);   
    }
}

public static class CurveBootstrapper
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
        Name = "d.Curve_Bootstrap",
        Description = "Bootstraps a single curve i.e. this is not a multi-curve bootstrapper.",
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
        DateTime baseDate =
            ExcelTableUtils.GetTableValue<DateTime>(curveParameters, "Value", "BaseDate", columnHeaderIndex);
        if (baseDate == default)
        {
            return CommonUtils.CurveParameterMissingErrorMessage(nameof(baseDate).ToUpper());
        }

        Settings.setEvaluationDate(baseDate);

        string? rateIndexName =
            ExcelTableUtils.GetTableValue<string>(curveParameters, "Value", "RateIndexName", columnHeaderIndex);
        if (rateIndexName is null && customRateIndex is null)
        {
            return CommonUtils.CurveParameterMissingErrorMessage(nameof(rateIndexName).ToUpper());
        }

        string? rateIndexTenor =
            ExcelTableUtils.GetTableValue<string>(curveParameters, "Value", "RateIndexTenor", columnHeaderIndex);
        if (rateIndexTenor is null && customRateIndex is null && rateIndexName != "FEDFUND")
        {
            return CommonUtils.CurveParameterMissingErrorMessage(nameof(rateIndexTenor).ToUpper());
        }

        string? interpolation =
            ExcelTableUtils.GetTableValue<string>(curveParameters, "Value", "Interpolation", columnHeaderIndex);
        if (interpolation == null)
        {
            return CommonUtils.CurveParameterMissingErrorMessage(nameof(rateIndexTenor).ToUpper());
        }
        
        string? discountCurveHandle =
            ExcelTableUtils.GetTableValue<string>(curveParameters, "Value", "DiscountCurveHandle", columnHeaderIndex);

        
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
            object[,] instruments = (object[,]) instrumentGroup;
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
            else if (string.Compare(instrumentType, "Interest Rate Swaps",
                         StringComparison.InvariantCultureIgnoreCase) == 0)
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
                        RelinkableHandle<YieldTermStructure>? discountCurve = null;
                        if (discountCurveHandle != null)
                        {
                            discountCurve.linkTo(CurveUtils.GetCurveObject(discountCurveHandle));
                        }
                        
                        rateHelpers.Add(
                            item: new SwapRateHelper(
                                rate: new Handle<Quote>(new SimpleQuote(rates[i])),
                                tenor: new Period(tenors[i]),
                                calendar: rateIndex.fixingCalendar(),
                                fixedFrequency: Frequency.Quarterly,
                                fixedConvention: rateIndex.businessDayConvention(),
                                fixedDayCount: rateIndex.dayCounter(),
                                iborIndex: rateIndex,
                                discount: discountCurve));
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
                            item: new OISRateHelper(2, new Period(tenors?[i]),
                                new Handle<Quote>(new SimpleQuote(rates?[i])), rateIndex as OvernightIndex));
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

        CurveDetails curveDetails = new(termStructure, rateIndex.dayCounter(), interpolation, new List<Date>(),
            new List<double>(), instrumentGroups);
        DataObjectController dataObjectController = DataObjectController.Instance;
        return dataObjectController.Add(handle, curveDetails);
    }


    [ExcelFunction(
        Name = "d.Curve_Get",
        Description = "Extracts and bootstraps a curve from the Omicron database.",
        Category = "∂Excel: Interest Rates")]
    public static string Get(
        string handle,
        string curveName,
        DateTime baseDate,
        string interpolation = "Exponential")
    {
        // Assume has deposits, FRAs, and Swaps
        // Could create more complicated abstract code for mapping from quotes to 2d tables but I would advise against this.
        string rateIndexName = "";
        string rateIndexTenor = "";
        switch (curveName.ToUpper())
        {
            case "ZAR-SWAP":
                rateIndexName = "JIBAR"; 
                rateIndexTenor = "3m";
                break;
            case "USD-OIS":
                rateIndexName = "FEDFUND";
                rateIndexTenor = "1d";
                break;
        }
        
        List<QuoteValue> quoteValues =
            OmicronUtils.GetSwapCurveQuotes(rateIndexName, null, 1, baseDate.ToString("yyyy-MM-dd"));

        object[,] curveParameters =
        {
            {"CurveUtils Parameters", ""},
            {"Parameter", "Value"},
            {"BaseDate", baseDate.ToOADate()},
            {"RateIndexName", rateIndexName},
            {"RateIndexTenor", rateIndexTenor},
            {"Interpolation", interpolation},
        };

        List<QuoteValue> deposits = quoteValues.Where(x => x.Type.GetType() == typeof(RateIndex)).ToList();
        deposits = deposits.OrderBy(x => ((RateIndex)x.Type).Tenor, new TenorComparer()).ToList();
        object[,] depositInstruments = new object[deposits.Count + 2, 4];
        depositInstruments[0, 0] = "Deposits";
        depositInstruments[1, 0] = "Tenors";
        depositInstruments[1, 1] = "RateIndex";
        depositInstruments[1, 2] = "Rates";
        depositInstruments[1, 3] = "Include";

        int row = 2;
        foreach (QuoteValue deposit in deposits)
        {
            depositInstruments[row, 0] = ((RateIndex) deposit.Type).Tenor.ToString();
            depositInstruments[row, 1] = ((RateIndex) deposit.Type).Name;
            depositInstruments[row, 2] = deposit.Value;
            depositInstruments[row, 3] = "TRUE";
            row++;
        }

        List<QuoteValue> fras = quoteValues.Where(x => x.Type.GetType() == typeof(Fra)).ToList();
        fras = fras.OrderBy(x => ((Fra)x.Type).Tenor, new TenorComparer()).ToList();
        object[,] fraInstruments = new object[fras.Count + 2, 4];
        row = 2;
        fraInstruments[0, 0] = "FRAs";
        fraInstruments[1, 0] = "FraTenors";
        fraInstruments[1, 1] = "RateIndex";
        fraInstruments[1, 2] = "Rates";
        fraInstruments[1, 3] = "Include";

        foreach (QuoteValue fra in fras)
        {
            // TODO: Ensure the amount is always in months.
            fraInstruments[row, 0] = $"{((Fra) fra.Type).Tenor.Amount}x{((Fra) fra.Type).Tenor.Amount + 3}";
            fraInstruments[row, 1] = ((Fra) fra.Type).ReferenceIndex.Name;
            fraInstruments[row, 2] = fra.Value;
            fraInstruments[row, 3] = "TRUE";
            row++;
        }

        List<QuoteValue> swaps = quoteValues.Where(x => x.Type.GetType() == typeof(InterestRateSwap)).ToList();
        swaps = swaps.OrderBy(x => ((InterestRateSwap)x.Type).Tenor, new TenorComparer()).ToList();
        object[,] swapInstruments = new object[swaps.Count + 2, 4];
        swapInstruments[0, 0] = "Interest Rate Swaps";
        swapInstruments[1, 0] = "Tenors";
        swapInstruments[1, 1] = "RateIndex";
        swapInstruments[1, 2] = "Rates";
        swapInstruments[1, 3] = "Include";

        row = 2;
        foreach (QuoteValue swap in swaps)
        {
            swapInstruments[row, 0] = ((InterestRateSwap) swap.Type).Tenor.ToString();
            swapInstruments[row, 1] = ((InterestRateSwap) swap.Type).ReferenceIndex.Name;
            swapInstruments[row, 2] = swap.Value;
            swapInstruments[row, 3] = "TRUE";
            row++;
        }
        
        List<QuoteValue> oiss = quoteValues.Where(x => x.Type.GetType() == typeof(Ois)).ToList();
        oiss = oiss.OrderBy(x => ((Ois)x.Type).Tenor, new TenorComparer()).ToList();
        object[,] oisInstruments = new object[oiss.Count + 2, 3];
        oisInstruments[0, 0] = "OISs";
        oisInstruments[1, 0] = "Tenors";
        oisInstruments[1, 1] = "Rates";
        oisInstruments[1, 2] = "Include";

        row = 2;
        foreach (QuoteValue ois in oiss)
        {
            oisInstruments[row, 0] = ((Ois)ois.Type).Tenor.ToString();
            oisInstruments[row, 1] = ois.Value;
            oisInstruments[row, 2] = "TRUE";
            row++;
        }

        List<object> instruments = new();

        if (deposits.Any())
        {
            instruments.Add(depositInstruments);
        }

        if (fras.Any())
        {
            instruments.Add(fraInstruments);
        }
        
        if (swaps.Any())
        {
            instruments.Add(swapInstruments);
        }
        
        if (oiss.Any())
        {
            instruments.Add(oisInstruments);
        }
        
        return Bootstrap(handle, curveParameters, null, instruments.ToArray());
    }
}
