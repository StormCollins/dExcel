namespace dExcel.FX;

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
