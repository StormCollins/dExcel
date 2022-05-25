namespace dExcel;

using ExcelDna.Integration;
using QLNet;

public static class HullWhite
{
    [ExcelFunction(
        Name = "d.HullWhite_Calibrate",
        Description = "Used to calibrate the Hull-White model on co-terminal swaption quotes.",
        Category = "∂Excel: Interest Rates")]
    public static string Calibrate(
        string parameterHandle, 
        string curveHandle, 
        object[,] swaptionMaturities, 
        object[,] swapLengths, 
        object[,] swaptionVols)
    {
        
        
        var discountCurve = Curves.Curve.GetDiscountCurve(curveHandle);
        var termStructure = new Handle<YieldTermStructure>(discountCurve);
        var jibar = new Jibar(new Period(3, TimeUnit.Months), termStructure);

        List<CalibrationHelper> swaptions = new();
        for (int i = 0; i < swaptionMaturities.GetLength(0); i++)
        {
            swaptions.Add(new SwaptionHelper(
                maturity: new Period((int)swaptionMaturities[i, 0], TimeUnit.Years),
                length: new Period((int)swapLengths[i, 0], TimeUnit.Years),
                volatility: new Handle<Quote>(new SimpleQuote((double)swaptionVols[i, 0])),
                index: jibar,
                fixedLegTenor: jibar.tenor(),
                fixedLegDayCounter: jibar.dayCounter(),
                floatingLegDayCounter: jibar.dayCounter(),
                termStructure,
                type: VolatilityType.ShiftedLognormal));
        }

        QLNet.HullWhite hullWhite = new QLNet.HullWhite(termStructure);
        for (int i = 0; i < swaptions.Count; i++)
        {
            swaptions[i].setPricingEngine(new JamshidianSwaptionEngine(hullWhite));
        }
        CalibrateModel(hullWhite, swaptions);
        var parameters = new Dictionary<string, double>()
        {
            ["ɑ"] = hullWhite.parameters()[0],
            ["σ"] = hullWhite.parameters()[1],
        };
        return DataObjectController.Add(parameterHandle, parameters);
    }
    
    static void CalibrateModel(ShortRateModel model,
        List<CalibrationHelper> helpers)
    {
        if (model == null)
            throw new ArgumentNullException("model");
        var om = new LevenbergMarquardt();
        model.calibrate(helpers, om,
            new EndCriteria(400, 100, 1.0e-8, 1.0e-8, 1.0e-8), new Constraint(),
            new List<double>());
    }

    public static object[,] GetHullWhiteParameters(string handle)
    {
        var parameters = (Dictionary<string, double>)DataObjectController.GetDataObject(handle);
        object[,] output = new object[3, 2];
        output[0, 0] = "Parameter";
        output[0, 1] = "Value";
        output[1, 0] = "ɑ";
        output[1, 1] = parameters["alpha"];
        output[2, 0] = "σ";
        output[2, 1] = parameters["sigma"];
        return output;
    }
}
