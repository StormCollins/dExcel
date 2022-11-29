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
    /// <summary>
    /// Calculates the historic volatility of an equity.
    /// </summary>
    /// <param name="valuationDate">The valuation date.</param>
    /// <param name="maturityDate">The maturity date.</param>
    /// <param name="businessDaysPerYear">The number of business days in a year.</param>
    /// <param name="dates">The dates for the corresponding stock prices.</param>
    /// <param name="prices">The stock prices for the corresponding dates.</param>
    /// <param name="weightingStyle">Set to 'Equal' or 'Exponential' for equally or exponentially weighted volatilities respectively.</param>
    /// <param name="lambda"></param>
    /// <returns></returns>
    [ExcelFunction(
        Name = "d.Equity_Volatility",
        Description = "Calculates the historic volatility of an equity.\n" +
                      "Deprecates AQS function: 'DT_Volatility'",
        Category = "∂Excel: Equities")]
    public static object Volatility(
        [ExcelArgument(
            Name = "Valuation Date",
            Description = "The valuation date.")]
        DateTime valuationDate,
        [ExcelArgument(
            Name = "Maturity Date",
            Description = "The maturity date.")]
        DateTime maturityDate,
        [ExcelArgument(
            Name = "Business Days in Year",
            Description = "The number of business days in a year.")]
        int businessDaysPerYear,
        [ExcelArgument(
            Name = "Weighting Style",
            Description = "Set to 'Equal' or 'Exponential' for equally or exponentially weighted volatilities respectively.")]
        string weightingStyle,
        [ExcelArgument(
            Name = "Lambda",
            Description = "EWMA Lambda parameter.")]
        double lambda,
        [ExcelArgument(
            Name = "Dates",
            Description = "The dates for the corresponding stock prices.")]
        object[,] dates,
        [ExcelArgument(
            Name = "Prices",
            Description = "Stock prices for the corresponding dates.")]
        object[,] prices)
    {
#if DEBUG
        CommonUtils.InFunctionWizard();
#endif
        if (dates.GetLength(0) != prices.GetLength(0))
        {
            return CommonUtils.DExcelErrorMessage("Date array and stock price array have different sizes.");
        }

        List<(DateTime date, double price)> datesAndPrices = new();
        for (int i = 0; i < dates.GetLength(0); i++)
        {
            datesAndPrices.Add((DateTime.FromOADate((double)dates[i, 0]), (double)prices[i, 0]));
        }

        List<(DateTime date, double price)> sortedDatesAndPrices = datesAndPrices.OrderBy(x => x.date).ToList();
        DateTime endDate = valuationDate;
        DateTime startDate = valuationDate - (maturityDate - valuationDate);

        // Determine the relevant time period over which volatility should be calculated.
        List<(DateTime date, double price)> sortedDatesAndPricesForVolCalculation = sortedDatesAndPrices.Where(x => startDate <= x.date && x.date <= endDate).ToList();
        List<double> returns = new List<double>();
        for (int i = 0; i < sortedDatesAndPricesForVolCalculation.Count - 1; i++)
        {
            returns.Add(Math.Log(sortedDatesAndPricesForVolCalculation[i].price / sortedDatesAndPricesForVolCalculation[i + 1].price));
        }

        double equallyWeightedVolatility = 0.0;
        double ewmaVolatility = 0.0;
        if (weightingStyle.ToUpper() == "EQUAL")
        {
            equallyWeightedVolatility = mns.Statistics.StandardDeviation(returns) * Math.Sqrt(businessDaysPerYear);
        }
        else if (weightingStyle.ToUpper() == "EXPONENTIAL")
        {
            List<double> squareReturns = returns.Select(x => x * x).ToList();
            double initialEwma = Math.Sqrt(squareReturns.Skip(Math.Max(0, squareReturns.Count() - 24)).Sum());
            List<double> ewmaSeries = Enumerable.Repeat(0.0, squareReturns.Count - 1).Append(initialEwma).ToList();
            int m = ewmaSeries.Count - 1;
            for (int i = ewmaSeries.Count - 1; i > 0; i--)
            {
                ewmaSeries[i] = Math.Sqrt(Math.Pow(ewmaSeries[i + 1], 2) * lambda + (1 - lambda) * squareReturns[i] * businessDaysPerYear);
            }
            ewmaVolatility = ewmaSeries[0];
        }
        else
        {
            return CommonUtils.DExcelErrorMessage($"Invalid weighting style: {weightingStyle}");
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
        [ExcelArgument(Name = "Long/Short", Description = "'Long' or 'Short' position.")]
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
                return CommonUtils.DExcelErrorMessage($"Invalid 'long'/'short' position: {longOrShort}");
        }

        if (spotPrice <= 0)
        {
            return CommonUtils.DExcelErrorMessage($"Spot price cannot be negative: {spotPrice}");
        }

        if (vol <= 0)
        {
            return CommonUtils.DExcelErrorMessage($"Volatility cannot be negative: {vol}");
        }

        if (dividendYield <= 0)
        {
            return CommonUtils.DExcelErrorMessage($"Dividend yield cannot be negative: {dividendYield}");
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

