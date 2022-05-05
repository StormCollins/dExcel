
namespace dExcel;

using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using mns = MathNet.Numerics.Statistics;

/// <summary>
/// A collection of utility functions for equities.
/// </summary>
public static class EquityUtils
{
    [ExcelFunction(
        Name = "d.Volatility",
        Description = "Calculates the historic volatility of an equity.\nDeprecates AQS function: ''",
        Category = "∂Excel: Equities")]
    public static object Volatility(
        [ExcelArgument(
            Name = "Dates and Prices",
            Description = "The two columned range containing the dates and prices.")]
        object[,] priceData,
        [ExcelArgument(
            Name = "Valuation Date",
            Description = "The valuation date.")]
        DateTime valDate,
        [ExcelArgument(
            Name = "Maturity Date", 
            Description = "The maturity date.")]
        DateTime maturityDate,
        [ExcelArgument(
            Name = "Equally Weighted",
            Description = "(Boolean) Set to 'True' to calculate the equally weighted vol.")]
        bool equallyWeighted,
        [ExcelArgument(
            Name = "Exp. Weighted",
            Description = "(Boolean) Set to 'True' to calculate the exponentially weighted vol.")]
        bool exponentiallyWeighted,
        [ExcelArgument(
            Name = "Business Day Count",
            Description = "The rolling number of business days in a year.")]
        int businessDaysPerYear,
        [ExcelArgument(
            Name = "Lambda",
            Description = "EWMA Lambda parameter.")]
        double lambda)
    {
#if DEBUG
        CommonUtils.InFunctionWizard();
#endif

        var datesAndPrices = new List<(DateTime date, double price)>();
        for (int i = 0; i < priceData.GetLength(0); i++)
        {
            datesAndPrices.Add((DateTime.FromOADate((double)priceData[i, 0]), (double)priceData[i, 1]));
        }

        var sortedDatesAndPrices = datesAndPrices.OrderBy(x => x.date).ToList();
        var endDate = valDate;
        var startDate = valDate - (maturityDate - valDate);

        // Determine the relevant time period over which volatility should be calculated
        var sortedDatesAndPricesForVolCalculation = sortedDatesAndPrices.Where(x => startDate <= x.date && x.date <= endDate).ToList();
        var returns = new List<double>();
        for (int i = 0; i < sortedDatesAndPricesForVolCalculation.Count - 1; i++)
        {
            returns.Add(Math.Log(sortedDatesAndPricesForVolCalculation[i].price / sortedDatesAndPricesForVolCalculation[i + 1].price));
        }

        var equallyWeightedVolatility = 0.0;
        if (equallyWeighted)
        {
            equallyWeightedVolatility = mns.Statistics.StandardDeviation(returns) * Math.Sqrt(businessDaysPerYear);
        }

        var ewmaVolatility = 0.0;
        if (exponentiallyWeighted)
        {
            var squareReturns = returns.Select(x => x * x).ToList();
            var initialEwma = Math.Sqrt(squareReturns.Skip(Math.Max(0, squareReturns.Count() - 24)).Sum());
            var ewmaSeries = Enumerable.Repeat(0.0, squareReturns.Count - 1).Append(initialEwma).ToList();
            var m = ewmaSeries.Count - 1;
            for (int i = ewmaSeries.Count - 1; i > 0; i--)
            {
                ewmaSeries[i] = Math.Sqrt(Math.Pow(ewmaSeries[i + 1], 2) * lambda + (1 - lambda) * squareReturns[i] * businessDaysPerYear);
            }
            ewmaVolatility = ewmaSeries[0];
        }

        if (equallyWeighted && exponentiallyWeighted)
        {
            return new double[1, 2] { { equallyWeightedVolatility, ewmaVolatility } };
        }

        return Math.Max(equallyWeightedVolatility, ewmaVolatility);
    }
}
