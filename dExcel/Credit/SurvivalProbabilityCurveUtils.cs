using dExcel.Dates;
using ExcelDna.Integration;
using QL = QuantLib;

namespace dExcel.Credit;

/// <summary>
/// A collection of utility functions for working with survival probability curves.
/// </summary>
public static class SurvivalProbabilityCurveUtils
{
    // TODO: Check if survival probabilities use Act360 only or if that's just for CDSs.
    /// <summary>
    /// Creates a survival probability curve from a set of survival probabilities.
    /// </summary>
    /// <param name="handle">The 'handle' or name used to refer to the object in memory.
    /// Each 'object' in a workbook must have a a unique handle.</param>
    /// <param name="datesRange">The Excel range containing the dates.</param>
    /// <param name="survivalProbabilitiesRange">The Excel range containing the survival probabilities.</param>
    /// <returns>A handle to survival probability curve object.</returns>
    [ExcelFunction(
        Name = "d.Credit_CreateSurvivalProbabilityCurve",
        Description = "Creates a survival probability curve object from dates and survival probabilities.",
        Category = "∂Excel: Credit")]
    public static string Create_SurvivalProbabilityCurve(
        [ExcelArgument(
            Name = "Handle",
            Description =
                "The 'handle' or name used to refer to the object in memory.\n" +
                "Each curve in a workbook must have a a unique handle.")]
        string handle,
        [ExcelArgument(
            Name = "Dates", 
            Description = "The dates for the corresponding discount factors.")]
        object[,] datesRange,
        [ExcelArgument(
            Name = "Survival Probabilities", 
            Description = "The survival probabilities for the corresponding dates.")]
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
