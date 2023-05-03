using ExcelDna.Integration;
using QL = QuantLib;

namespace dExcel;

public static class HullWhite
{
    /// <summary>
    /// Used to calibrate the Hull-White model on swaption quotes.
    /// </summary>
    /// <param name="parameterHandle">The handle for the calibration object returned by this function.</param>
    /// <param name="curveHandle">The handle to the discount curve.</param>
    /// <param name="swaptionData">The swaption data.</param>
    /// <returns>A handle to calibration object containing the calibration parameters.</returns>
    [ExcelFunction(
        Name = "d.HullWhite_Calibrate",
        Description = "Used to calibrate the Hull-White model on swaption quotes.",
        Category = "∂Excel: Interest Rates")]
    public static string Calibrate(
        string parameterHandle, 
        string curveHandle, 
        object[,] swaptionData)
    {
        int swaptionMaturitiesIndex = 0;
        int swapLengthsIndex = 0;
        int swaptionVolsIndex = 0;
        int volTypesIndex = 0;
        
        for (int j = 0; j < swaptionData.GetLength(1); j++)
        {
            switch (swaptionData[0, j].ToString()?.ToUpper())
            {
                case "SWAPTION MATURITIES":
                    swaptionMaturitiesIndex = j;
                    break;
                case "SWAP LENGTHS":
                    swapLengthsIndex = j;
                    break;
                case "SWAPTION VOLS":
                    swaptionVolsIndex = j;
                    break;
                case "VOL TYPE":
                    volTypesIndex = j;
                    break;
                default:
                    break;
            }         
        }

        List<int> swaptionMaturities = new();
        for (var i = 1; i < swaptionData.GetLength(1); i++)
        {
            swaptionMaturities.Add(int.Parse(swaptionData[i, swaptionMaturitiesIndex].ToString()));
        }
        
        List<int> swapLengths = new();
        for (var i = 1; i < swaptionData.GetLength(1); i++)
        {
            swapLengths.Add(int.Parse(swaptionData[i, swapLengthsIndex].ToString()));
        }
        
        List<double> swaptionVols = new();
        for (var i = 1; i < swaptionData.GetLength(1); i++)
        {
            swaptionVols.Add(double.Parse(swaptionData[i, swaptionVolsIndex].ToString()));
        }
        
        List<string?> volTypes = new();
        for (var i = 1; i < swaptionData.GetLength(1); i++)
        {
            volTypes.Add(swaptionData[i, volTypesIndex].ToString());
        }
        
        // var discountCurve = CurveUtils.GetDiscountCurve(curveHandle);
        // var termStructure = new Handle<YieldTermStructure>(discountCurve);
        // var jibar = new Jibar(new Period(3, TimeUnit.Months), termStructure);
        //
        // List<CalibrationHelper> swaptions = new();
        // for (int i = 0; i < swaptionMaturities.Count; i++)
        // {
        //     swaptions.Add(new SwaptionHelper(
        //         maturity: new Period(swaptionMaturities[i], TimeUnit.Years),
        //         length: new Period(swapLengths[i], TimeUnit.Years),
        //         volatility: new Handle<Quote>(new SimpleQuote((double)swaptionVols[i])),
        //         index: jibar,
        //         fixedLegTenor: jibar.tenor(),
        //         fixedLegDayCounter: jibar.dayCounter(),
        //         floatingLegDayCounter: jibar.dayCounter(),
        //         termStructure,
        //         type: VolatilityType.ShiftedLognormal));
        // }
        //
        // QLNet.HullWhite hullWhite = new(termStructure);
        // for (int i = 0; i < swaptions.Count; i++)
        // {
        //     swaptions[i].setPricingEngine(new JamshidianSwaptionEngine(hullWhite));
        // }
        // CalibrateModel(hullWhite, swaptions);
        // var parameters = new Dictionary<string, double>()
        // {
        //     ["alpha"] = hullWhite.parameters()[0],
        //     ["sigma"] = hullWhite.parameters()[1],
        // };
        //
        // DataObjectController dataObjectController = DataObjectController.Instance;
        // return dataObjectController.Add(parameterHandle, parameters);
        return "";
    }
    
    /// <summary>
    /// Calibrates the short rate model to the given calibration helpers.
    /// </summary>
    /// <param name="model">The short rate model to calibrate.</param>
    /// <param name="helpers">The calibration instruments.</param>
    /// <exception cref="ArgumentNullException">Thrown if the model is null.</exception>
    private static void CalibrateModel(QL.ShortRateModel model, QL.CalibrationHelperVector helpers)
    {
        if (model == null)
        {
            throw new ArgumentNullException($"Missing short rate model: {nameof(model)} cannot be null.");
        }

        model.calibrate(
            helpers,
            new QL.LevenbergMarquardt(),
            new QL.EndCriteria(400, 100, 1.0e-8, 1.0e-8, 1.0e-8));
    }

    /// <summary>
    /// Extracts the Hull-White model parameters from the relevant calibration object handle.
    /// </summary>
    /// <param name="handle">The calibration object handle.</param>
    /// <returns>A grid containing the Hull-White parameters.</returns>
    [ExcelFunction(
        Name = "d.HullWhite_GetParameters",
        Description = "Used to extract parameters from a calibration object.",
        Category = "∂Excel: Interest Rates")]
    public static object[,] GetHullWhiteParameters(string handle)
    {
        DataObjectController dataObjectController = DataObjectController.Instance;
        Dictionary<string, double> parameters = (Dictionary<string, double>)dataObjectController.GetDataObject(handle);
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
