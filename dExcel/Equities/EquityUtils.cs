using ExcelDna.Integration;
using dExcel.Dates;
using dExcel.Utilities;
using mnd = MathNet.Numerics.Distributions;
using mns = MathNet.Numerics.Statistics;
using QL = QuantLib;

namespace dExcel.Equities;

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
        List<double> returns = new();
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
        Name = "d.Equity_CreateEuropeanOption",
        Description = "Create an European equity option.",
        Category = "∂Excel: Equities")]
    public static string CreateEuropeanOption(
        string handle,
        double spot,
        double strike, 
        double riskFreeRate,
        double volatility,
        DateTime tradeDate,
        DateTime maturityDate,
        string dayCountConvention,
        string putOrCall)
    {
        QL.Option.Type optionType;
        
        if (string.Compare(putOrCall, "Put", StringComparison.OrdinalIgnoreCase) == 0)
        {
            optionType = QL.Option.Type.Put; 
        }
        else
        {
            optionType = QL.Option.Type.Call;
        }
        
        QL.PlainVanillaPayoff payoff = new(optionType, strike);
        QL.EuropeanExercise exercise = new(maturityDate.ToQuantLibDate()); 
        QL.VanillaOption vanillaOption = new(payoff, exercise);
        QL.QuoteHandle spotHandle = new(new QL.SimpleQuote(spot));
        
        CommonUtils.TryParseDayCountConvention(dayCountConvention, out QL.DayCounter? dayCounter, out string errorMessage);

        QL.FlatForward interestRateCurve =
            new(tradeDate.ToQuantLibDate(),
                new QL.QuoteHandle(new QL.SimpleQuote(riskFreeRate)),
                dayCounter);
        
        QL.FlatForward dividendCurve =
            new(tradeDate.ToQuantLibDate(),
                new QL.QuoteHandle(new QL.SimpleQuote(0)),
                dayCounter);
        
        QL.BlackConstantVol constantVol = new(tradeDate.ToQuantLibDate(), new QL.SouthAfrica(), volatility, dayCounter);
        
        QL.BlackScholesMertonProcess blackScholesMertonProcess = 
            new(spotHandle, 
                new QL.YieldTermStructureHandle(dividendCurve),
                new QL.YieldTermStructureHandle(interestRateCurve),
                new QL.BlackVolTermStructureHandle(constantVol));
        
        vanillaOption.setPricingEngine(new QL.AnalyticEuropeanEngine(blackScholesMertonProcess));
        var dataObjectController = DataObjectController.Instance;
        return dataObjectController.Add(handle, vanillaOption);
    }

    [ExcelFunction(
        Name = "d.Equity_GetDetails",
        Description = "Get price of equity option.",
        Category = "∂Excel:Equities")]
    public static object GetOptionDetails(string handle)
    {
        var dataObjectController = DataObjectController.Instance; 
        QL.VanillaOption option = (QL.VanillaOption)dataObjectController.GetDataObject(handle);
        object[,] output = new object[,]
        {
            {"Price", option.NPV()},
            {"Delta", option.delta()},
            {"Vega", option.vega()},
        };
        
        return output;
    }


    [ExcelFunction(
        Name = "d.Portfolio_CreatePortfolio",
        Description = "Create portfolio.",
        Category = "∂Excel: Equities")]
    public static string CreatePortfolio(string handle, params object[] handles)
    {
        var dataObjectController = DataObjectController.Instance;
        List<string> handlesList = new();
        foreach (var currentHandle in handles)
        {
            handlesList.Add(currentHandle.ToString());
        }
        
        return dataObjectController.Add(handle, handlesList); 
    }
}
