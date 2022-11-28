namespace dExcel.CreditUtils;

using System;
using System.Collections.Generic;
using ExcelDna.Integration;
using QLNet;

public static class CreditUtils
{
    // TODO: Check if survival probabilities use Act360 only or if that's just for CDSs.
    [ExcelFunction(
        Name = "d.Credit_CreateSurvivalProbabilityCurve",
        Description = "Creates a survival probability curve from dates and survival probabilities.",
        Category = "∂Excel: Credit")]
    public static string Create_SurvivalProbabilityCurve(
        string handle,
        object[,] datesRange,
        object[,] survivalProbabilitiesRange)
    {
        List<Date> dates = new();
        List<double> survivalProbabilities = new();
        for (int i = 0; i < datesRange.GetLength(0); i++)
        {
            dates.Add((Date)DateTime.FromOADate((double)datesRange[i, 0]));
            survivalProbabilities.Add((double)survivalProbabilitiesRange[i, 0]);
        }
        InterpolatedSurvivalProbabilityCurve<LogLinear> curve =
            new(dates, survivalProbabilities, new Actual360());
        return DataObjectController.Add(handle, curve);
    }
}
