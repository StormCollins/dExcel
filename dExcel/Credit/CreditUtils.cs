using ExcelDna.Integration;
using QL = QuantLib;

namespace dExcel.CreditUtils;

using Dates;

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
        QL.DateVector dates = new();
        QL.DoubleVector survivalProbabilities = new();
        for (int i = 0; i < datesRange.GetLength(0); i++)
        {
            dates.Add(DateTime.FromOADate((double)datesRange[i, 0]).ToQuantLibDate());
            survivalProbabilities.Add((double)survivalProbabilitiesRange[i, 0]);
        }
        
        QL.SurvivalProbabilityCurve curve = new(dates, survivalProbabilities, new QL.Actual360());

        DataObjectController dataObjectController = DataObjectController.Instance;
        return dataObjectController.Add(handle, curve);
    }
}
