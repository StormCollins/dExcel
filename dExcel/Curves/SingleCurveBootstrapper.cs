namespace dExcel.Curves;

using ExcelDna.Integration;
using ExcelUtils;
using QLNet;

public static class SingleCurveBootstrapper
{
    [ExcelFunction(
        Name = "d.Curve_SingleCurveBootstrap",
        Description = "Bootstraps a single curve i.e. this is not a multi-curve bootstrapper.",
        Category = "∂Excel: Interest Rates")]
    public static string Bootstrap(string handle, DateTime baseDate, params object[] instrumentGroups)
    {
        Settings.setEvaluationDate(baseDate);
        List<RateHelper> rateHelpers = new();
        IborIndex rateIndex = null;
        
        foreach (var instrumentGroup in instrumentGroups)
        {
            var instruments = (object[,])instrumentGroup;
            
            // TODO: Make this case insensitive.
            List<string> tenors = ExcelTable.GetColumn<string>(instruments, "Tenors");
            List<DateTime> startDates = ExcelTable.GetColumn<DateTime>(instruments, "StartDates");
            List<DateTime> endDates = ExcelTable.GetColumn<DateTime>(instruments, "EndDates");
            List<string> rateIndices = ExcelTable.GetColumn<string>(instruments, "RateIndex");
            List<double> rates = ExcelTable.GetColumn<double>(instruments, "Rates");
            List<bool> include = ExcelTable.GetColumn<bool>(instruments, "Include");

            var instrumentCount = include.Count;


            var index = rateIndices[0];
            rateIndex =
                index switch
                {
                    "EURIBOR" => new Euribor(new Period("3m")),
                    "FEDFUND" => new FedFunds(),
                    "JIBAR" => new Jibar(new Period("3m")),
                    "USD-LIBOR" => new USDLibor(new Period("3m")),
                };

            string? instrumentType = ExcelTable.GetTableLabel(instruments);
            
            if (string.Compare(instrumentType, "Deposits", StringComparison.InvariantCultureIgnoreCase) == 0)
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
                                periodToStart: new Period(tenors[i]),
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
            ["Curve.Object"] = termStructure,
        };
        
        return DataObjectController.Add(handle, curveDetails);
    }
}
