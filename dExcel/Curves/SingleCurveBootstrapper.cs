namespace dExcel.Curves;

using ExcelDna.Integration;
using QLNet;

public static class SingleCurveBootstrapper
{
    [ExcelFunction(
        Name = "d.Curve_SingleCurveBootstrap",
        Description = "Bootstraps a single curve i.e. this is not a multi-curve bootstrapper.",
        Category = "∂Excel: Interest Rates")]
    public static string Bootstrap(string handle, DateTime baseDate, params object[] instrumentGroups)
    {
        List<RateHelper> rateHelpers = new();
        IborIndex rateIndex = null;
        
        for (int k = 0; k < instrumentGroups.Length; k++)
        {
            var instruments = (object[,])instrumentGroups[k];
            
            List<string> columnTitles
                = Enumerable
                    .Range(0, instruments.GetLength(1))
                    .Select(j => instruments[1, j])
                    .Cast<string>()
                    .ToList();

            // TODO: Make this case insensitive.
            List<string> tenors 
                = Enumerable
                    .Range(2, instruments.GetLength(0) - 2)
                    .Select(i => instruments[i, columnTitles.IndexOf("Tenors")])
                    .Cast<string>()
                    .ToList();
        
            List<string> rateIndices 
                = Enumerable
                    .Range(2, instruments.GetLength(0) - 2)
                    .Select(i => instruments[i, columnTitles.IndexOf("RateIndex")])
                    .Cast<string>()
                    .ToList();
            
            List<double> rates
                = Enumerable
                    .Range(2, instruments.GetLength(0) - 2)
                    .Select(i => instruments[i, columnTitles.IndexOf("Rates")])
                    .Cast<double>()
                    .ToList();
            
            var index = rateIndices[0];
            rateIndex =
                index switch
                {
                    "EURIBOR" => new Euribor(new Period("3m")),
                    "JIBAR" => new Jibar(new Period("3m")),
                    "USD-LIBOR" => new USDLibor(new Period("3m")),
                };


            string? instrumentType = instruments[0, 0].ToString();
            
            if (string.Compare(instrumentType, "Deposits", StringComparison.InvariantCultureIgnoreCase) == 0)
            {
                for (int i = 0; i < tenors.Count; i++)
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
            else if (string.Compare(instrumentType, "FRAs", StringComparison.InvariantCultureIgnoreCase) == 0)
            {
                for (int i = 0; i < tenors.Count; i++)
                {
                    rateHelpers.Add(
                        new FraRateHelper(
                            rate: new Handle<Quote>(new SimpleQuote(rates[i])),
                            periodToStart: new Period(tenors[i]),
                            iborIndex: rateIndex));
                }
            }
            else if (string.Compare(instrumentType, "Interest Rate Swaps", StringComparison.InvariantCultureIgnoreCase) == 0)
            {
                for (int i = 0; i < tenors.Count; i++)
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
            ["Curve"] = termStructure,
        };
        
        return DataObjectController.Add(handle, curveDetails);
    }
}
