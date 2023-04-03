namespace dExcel.Curves;

using ExcelUtils;
using ExcelDna.Integration;
using InterestRates;
using Omicron;
using OmicronUtils;
using QLNet;

public static class CurveBootstrapper
{
    [ExcelFunction(
        Name = "d.Curve_CurveBootstrap",
        Description = "Bootstraps a single curve i.e. this is not a multi-curve bootstrapper.",
        Category = "∂Excel: Interest Rates")]
    public static string Bootstrap(
        string handle,
        DateTime baseDate,
        params object[] instrumentGroups)
    {
        List<RateHelper> rateHelpers = new();
        IborIndex? rateIndex = null;

        foreach (object instrumentGroup in instrumentGroups)
        {
            object[,] instruments = (object[,]) instrumentGroup;

            List<string>? tenors = ExcelTableUtils.GetColumn<string>(instruments, "Tenors");
            List<string>? rateIndices = ExcelTableUtils.GetColumn<string>(instruments, "RateIndex");
            List<double>? rates = ExcelTableUtils.GetColumn<double>(instruments, "Rates");
            List<bool>? include = ExcelTableUtils.GetColumn<bool>(instruments, "Include");
            string index = rateIndices?[0];

            rateIndex =
                index switch
                {
                    "EURIBOR" => new Euribor(new Period("3m")),
                    "JIBAR" => new Jibar(new Period("3m")),
                    "USD-LIBOR" => new USDLibor(new Period("3m")),
                };

            string? instrumentType = ExcelTableUtils.GetTableLabel(instruments);

            if (string.Compare(instrumentType, "Deposits", StringComparison.InvariantCultureIgnoreCase) == 0)
            {
                for (int i = 0; i < tenors.Count; i++)
                {
                    if (include[i])
                    {
                        rateHelpers.Add(
                            new DepositRateHelper(
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
                for (int i = 0; i < tenors.Count; i++)
                {
                    if (include[i])
                    {
                        rateHelpers.Add(
                            new FraRateHelper(
                                rate: new Handle<Quote>(new SimpleQuote(rates[i])),
                                periodToStart: new Period(tenors[i]),
                                iborIndex: rateIndex));
                    }
                }
            }
            else if (string.Compare(instrumentType, "Interest Rate Swaps",
                         StringComparison.InvariantCultureIgnoreCase) == 0)
            {
                for (int i = 0; i < tenors.Count; i++)
                {
                    if (include[i])
                    {
                        SwapRateHelper x = new(
                            rate: new Handle<Quote>(new SimpleQuote(rates[i])),
                            tenor: new Period(tenors[i]),
                            calendar: rateIndex.fixingCalendar(),
                            fixedFrequency: Frequency.Quarterly,
                            fixedConvention: rateIndex.businessDayConvention(),
                            fixedDayCount: rateIndex.dayCounter(),
                            iborIndex: rateIndex);

                        rateHelpers.Add(
                            new SwapRateHelper(
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
        }

        YieldTermStructure termStructure =
            new PiecewiseYieldCurve<Discount, LogLinear>(
                new Date(baseDate),
                rateHelpers,
                rateIndex.dayCounter(),
                new List<Handle<Quote>>(),
                new List<Date>(),
                1.0e-20);

        Dictionary<string, object> curveDetails = new()
        {
            ["CurveUtils.Object"] = termStructure,
        };

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
        
        return SingleCurveBootstrapper.Bootstrap(handle, curveParameters, null, instruments.ToArray());
    }
}
