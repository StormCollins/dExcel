using dExcel.ExcelUtils;
using dExcel.InterestRates;
using ExcelDna.Integration;
using mnd = MathNet.Numerics.Distributions;
using QL = QuantLib;

namespace dExcel.FX;

/// <summary>
/// A collection of utility functions for working with FX.
/// </summary>
public static class FxUtils
{
    /// <summary>
    /// 
    /// </summary>
    /// <param name="volsRange"></param>
    /// <param name="deltasRange"></param>
    /// <param name="optionMaturitiesRange"></param>
    /// <param name="domesticCurveHandle"></param>
    /// <param name="foreignCurveHandle"></param>
    /// <returns></returns>
    [ExcelFunction(
        Name = "d.FX_ConvertDeltaToMoneynessVolSurface",
        Description = "Convert a delta vol surface to a moneyness vol surface.",
        Category = "∂Excel: FX")]
    public static object ConvertDeltaToMoneynessVolSurface(
        object[,] volsRange,
        object[,] deltasRange,
        object[,] optionMaturitiesRange,
        string domesticCurveHandle,
        string foreignCurveHandle)
    {
        List<double> deltas = ExcelArrayUtils.ConvertExcelRangeToList<double>(deltasRange);
        List<object> optionMaturities = 
            ExcelArrayUtils.ConvertExcelRangeToList<double>(optionMaturitiesRange).Cast<object>().ToList();
        List<(double moneyness, double optionMaturity, double vol)> moneynessSurface = new();
        List<double> moneynesses = new();

        List<double> domesticRates = 
            ExcelArrayUtils.ConvertExcelRangeToList<double>(
                (object[,])CurveUtils.GetDiscountFactors(domesticCurveHandle, optionMaturities.ToArray()));
        
        List<double> foreignRates = 
            ExcelArrayUtils.ConvertExcelRangeToList<double>(
                (object[,])CurveUtils.GetDiscountFactors(domesticCurveHandle, optionMaturities.ToArray()));
        
        for (int i = 0; i < optionMaturities.Count; i++)
        {
            for (int j = 0; j < deltas.Count; j++)
            {
                double moneyness = 
                    Math.Exp(
                        mnd.Normal.InvCDF(0, 1, deltas[i]) * 
                        (double)volsRange[i, j] * Math.Sqrt((double)optionMaturities[i]) -
                        (domesticRates[i] - foreignRates[i] + 0.5 * (double)volsRange[i, j] * (double)volsRange[i, j]) * (double)optionMaturities[i]);

                moneynesses.Add(Math.Round(moneyness, 3));
                moneynessSurface.Add((Math.Round(moneyness, 3) , (double)optionMaturities[i], (double)volsRange[i, j]));
            } 
        }

        moneynesses = moneynesses.Distinct().ToList();
        moneynesses.Sort();

        object[,] output = new object[optionMaturities.Count + 1, moneynesses.Count + 1];

        foreach ((double moneyness, double optionMaturity, double vol) in moneynessSurface)
        {
            int moneynessIndex = moneynesses.IndexOf(moneyness);
            int optionMaturityIndex = optionMaturities.IndexOf(optionMaturity);
            output[optionMaturityIndex + 1, moneynessIndex + 1] = vol;
        }

        for (int j = 0; j < moneynesses.Count; j++)
        {
            output[0, j + 1] = moneynesses[j];
        }

        for (int i = 0; i < optionMaturities.Count; i++)
        {
            output[i + 1, 0] = optionMaturities[i];
        }
        
        return output;
    }

    /// <summary>
    /// Calculates the delta of an FX option (call or put) using the formula:
    ///
    /// Delta = e^(-r_f * T) Phi(d_1) for a call.
    ///
    /// Delta = e^(-r_f * T) * [Phi(d_1) - 1] for a put.
    ///
    /// Where d_1 = (log(S_0/K) + (r_d - r_f + 0.5 * vol^2) * T)/ (vol * sqrt(T))
    ///
    /// See: https://blog.quantinsti.com/basics-options-trading/
    /// </summary>
    /// <returns>The delta of an FX option.</returns>
    public static double CalculateDelta(
        double spot,
        double strike,
        double domesticRate,
        double foreignRate,
        double vol,
        double optionMaturity,
        string optionType)
    {
        double d1 =
            (Math.Log(spot / strike) + (domesticRate - foreignRate + 0.5 * vol * vol) * optionMaturity) /
            (vol * Math.Sqrt(optionMaturity));

        if (optionType.ToUpper() == "C" || optionType.ToUpper() == "CALL")
        {
            return Math.Exp(-1 * foreignRate * optionMaturity) * mnd.Normal.CDF(0, 1, d1);
        }

        return Math.Exp(-1 * foreignRate * optionMaturity) * (mnd.Normal.CDF(0, 1, d1) - 1);
    }
}
