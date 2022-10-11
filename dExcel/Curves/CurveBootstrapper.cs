namespace dExcel.Curves;

using ExcelUtils;
using ExcelDna.Integration;
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
            object[,] instruments = (object[,])instrumentGroup;

            List<string>? tenors = ExcelTable.GetColumn<string>(instruments, "Tenors");
            List<string>? rateIndices = ExcelTable.GetColumn<string>(instruments, "RateIndex");
            List<double>? rates = ExcelTable.GetColumn<double>(instruments, "Rates");
            List<bool>? include = ExcelTable.GetColumn<bool>(instruments, "Include");
            string index = rateIndices?[0];
            
            rateIndex =
                index switch
                {
                    "EURIBOR" => new Euribor(new Period("3m")),
                    "JIBAR" => new Jibar(new Period("3m")),
                    "USD-LIBOR" => new USDLibor(new Period("3m")),
                };

            string? instrumentType = ExcelTable.GetTableLabel(instruments);

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
            else if (string.Compare(instrumentType, "Interest Rate Swaps", StringComparison.InvariantCultureIgnoreCase) == 0)
            {
                for (int i = 0; i < tenors.Count; i++)
                {
                    if (include[i])
                    {
                        var x = new SwapRateHelper(
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
            ["Curve.Object"] = termStructure,
        };

        return DataObjectController.Add(handle, curveDetails);
    }
}
