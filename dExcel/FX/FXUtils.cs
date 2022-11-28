namespace dExcel.FX;

using QLNet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

public static class FxUtils
{
    public static object GarmanKohlhagen(
        DateTime valuationDate,
        DateTime maturityDate,
        double spotFxRate,
        double strike,
        double domesticInterestRate,
        double foreignInterestRate,
        double volatility,
        string optionType)
    {
        return 0;
    }
}
