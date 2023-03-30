namespace dExcel.FX;

using dExcel.ExcelUtils;
using ExcelDna.Integration;
using mnd = MathNet.Numerics.Distributions;
using QLNet;
using Utilities;

public static class FxUtils
{
    public static double GetStrikeVolFromDeltaSurface(object[,] deltaSurface, double spot, double optionStrike, double domesticDf, double foreignDf)
    {
        return 0;
    }

    [ExcelFunction(
        Name = "d.FX_ConvertDeltaToMoneynessVolSurface",
        Description = "Convert a delta vol surface to a moneyness vol surface.",
        Category = "∂Excel: FX")]
    public static object ConvertDeltaToMoneynessVolSurface(
        object[,] volsRange,
        object[,] deltasRange,
        object[,] optionMaturitiesRange,
        double domesticRate,
        double foreignRate)
    {
        List<double> deltas = ExcelArrayUtils.ConvertExcelRangeToList<double>(deltasRange);
        List<double> optionMaturities = ExcelArrayUtils.ConvertExcelRangeToList<double>(optionMaturitiesRange);
        List<(double moneyness, double optionMaturity, double vol)> moneynessSurface = new();
        List<double> moneynesses = new();

        for (int i = 0; i < optionMaturities.Count; i++)
        {
            for (int j = 0; j < deltas.Count; j++)
            {
                double moneyness = 
                    Math.Exp(
                        deltas[i] * (double)volsRange[i, j] * Math.Sqrt(optionMaturities[i]) -
                        (domesticRate - foreignRate + 0.5 * (double)volsRange[i, j] * (double)volsRange[i, j]) * optionMaturities[i]);

                moneynesses.Add(Math.Round(moneyness, 3));
                moneynessSurface.Add((Math.Round(moneyness, 3) , optionMaturities[i], (double)volsRange[i, j]));
            } 
        }

        moneynesses.Sort();

        object[,] output = new object[moneynesses.Count + 1, optionMaturities.Count + 1];

        foreach ((double moneyness, double optionMaturity, double vol) in moneynessSurface)
        {
            int moneynessIndex = moneynesses.IndexOf(moneyness);
            int optionMaturityIndex = optionMaturities.IndexOf(optionMaturity);
            output[optionMaturityIndex, moneynessIndex] = vol;
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
