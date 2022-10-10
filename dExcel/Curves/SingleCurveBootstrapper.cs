namespace dExcel.Curves;

using ExcelDna.Integration;
using ExcelUtils;
using QLNet;

public static class SingleCurveBootstrapper
{
    [ExcelFunction(
        Name = "d.Curve_SingleCurveBootstrap",
        Description = "Bootstraps a single curve i.e. this is not a multi-curve bootstrapper.\n" +
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
            return "#Error: Please provide a base date in the curve parameters.";
        }
        
        Settings.setEvaluationDate(baseDate);

        string? index = ExcelTable.GetTableValue<string>(curveParameters, "Value", "RateIndex");
        if (index is null && customRateIndex is null)
        {
            return "#Error: Please provide a rate index in the curve parameters.";
        }

        string? indexTenor = ExcelTable.GetTableValue<string>(curveParameters, "Value", "RateIndexTenor");
        if (indexTenor is null && customRateIndex is null && index != "FEDFUND")
        {
            return $"#Error: Please provide a rate index tenor in the curve parameters.";
        }

        IborIndex? rateIndex = null;

        if (index is not null)
        {
            rateIndex =
                index switch
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
            return $"#Error: Unsupported rate index: {index}";
        }

        List<RateHelper> rateHelpers = new();

        foreach (var instrumentGroup in instrumentGroups)
        {
            var instruments = (object[,])instrumentGroup;

            // TODO: Make this case insensitive.
            List<string>? tenors = ExcelTable.GetColumn<string>(instruments, "Tenors");
            List<string>? fraTenors = ExcelTable.GetColumn<string>(instruments, "FraTenors");
            List<DateTime>? startDates = ExcelTable.GetColumn<DateTime>(instruments, "StartDates");
            List<DateTime>? endDates = ExcelTable.GetColumn<DateTime>(instruments, "EndDates");
            //List<string>? rateIndices = ExcelTable.GetColumn<string>(instruments, "RateIndex");
            List<double>? rates = ExcelTable.GetColumn<double>(instruments, "Rates");
            List<bool>? include = ExcelTable.GetColumn<bool>(instruments, "Include");
            string? instrumentType = ExcelTable.GetTableLabel(instruments);

            var instrumentCount = include.Count;
            
            if (string.Equals(instrumentType, "Deposits", StringComparison.OrdinalIgnoreCase))
            {
                for (int i = 0; i < instrumentCount; i++)
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
                for (int i = 0; i < instrumentCount; i++)
                {
                    if (include[i])
                    {
                        rateHelpers.Add(
                            new FraRateHelper(
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
                    if (include[i])
                    {
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
            else if (string.Compare(instrumentType, "OISs", StringComparison.InvariantCultureIgnoreCase) == 0)
            {
                for (int i = 0; i < instrumentCount; i++)
                {
                    if (include[i])
                    {
                        rateHelpers.Add(
                            new DatedOISRateHelper(
                                startDate: baseDate, 
                                endDate: endDates[i],
                                new Handle<Quote>(new SimpleQuote(rates[i])),
                                (OvernightIndex)rateIndex));
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
            // TODO: Add instruments used in bootstrapping.
            // TODO: Get out anchor dates.
            ["Curve.Object"] = termStructure,
        };
        
        return DataObjectController.Add(handle, curveDetails);
    }
}
