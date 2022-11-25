using System.Data.Common;

namespace dExcel.EquityUtils;

using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using mnd = MathNet.Numerics.Distributions;
using mns = MathNet.Numerics.Statistics;

/// <summary>
/// A collection of utility functions for equities.
/// </summary>
public static class EquityUtils
{
    [ExcelFunction(
        Name = "d.Equity_Volatility",
        Description = "Calculates the historic volatility of an equity.\nDeprecates AQS function: 'DT_Volatility'",
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

        // Determine the relevant time period over which volatility should be calculated.
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

   [ExcelFunction(
       Name = "d.Equity_BlackScholes",
       Description = "Black-Scholes option pricer. \nDeprecates AQS function: 'BS'",
       Category = "∂Excel: Equities")]
    public static object BlackScholes(
        [ExcelArgument(Name = "Option Type", Description = "'Call'/'C' or 'Put'/'P'.")]
        string optionType,
        [ExcelArgument(Name = "Long/Short", Description = "'Long' or 'Short'.")]
        string longOrShort,
        [ExcelArgument(Name = "S", Description = "Current stock price.")]
        double spotPrice,
        [ExcelArgument(Name = "K", Description = "Strike.")]
        double strike,
        [ExcelArgument(Name = "r", Description = "Risk free (NACC) rate. Only required for discounting.")]
        double rate,
        [ExcelArgument(Name = "q", Description = "Dividend Yield (NACC).")]
        double dividendYield,
        [ExcelArgument(Name = "T", Description = "Time to maturity.")]
        double timeToMaturity,
        [ExcelArgument(Name = "σ", Description = "Volatility.")]
        double vol)
    {
#if DEBUG
        CommonUtils.InFunctionWizard();
#endif 
        int sign;

        switch (optionType.ToUpper())
        {
            case "C":
            case "CALL":
                sign = 1;
                break;
            case "P":
            case "PUT":
                sign = -1;
                break;
            default:
                return CommonUtils.DExcelErrorMessage($"Invalid option type: {optionType}");
        }

        int longOrShortDirection;
        switch (longOrShort.ToUpper())
        {
            case "LONG":
                longOrShortDirection = 1;
                break;
            case "SHORT":
                longOrShortDirection = -1;
                break;
            default:
                return CommonUtils.DExcelErrorMessage($"Invalid 'long'/'short' direction: {longOrShort}");
        }

        double d1 = (Math.Log(spotPrice / strike) + (rate-dividendYield+Math.Pow(vol, 2)/2) * timeToMaturity) / (vol * Math.Sqrt(timeToMaturity));
        double d2 = d1 - vol * Math.Sqrt(timeToMaturity);

        if (timeToMaturity <= 0)
        {
            return Math.Exp(-(rate-dividendYield) * timeToMaturity) * Math.Max(0, sign * (spotPrice - strike));
        }

        return longOrShortDirection *
               (sign * (spotPrice * Math.Exp(-dividendYield * timeToMaturity) * mnd.Normal.CDF(0, 1, sign * d1) -
                        strike * Math.Exp(-rate * timeToMaturity) * mnd.Normal.CDF(0, 1, sign * d2)));
    }
}

