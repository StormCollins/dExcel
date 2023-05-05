using ExcelDna.Integration;
using dExcel.Dates;
using dExcel.Utilities;
using mnd = MathNet.Numerics.Distributions;
using mns = MathNet.Numerics.Statistics;
using QL = QuantLib;

namespace dExcel.Equities;

using static CommonUtils;

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
        InFunctionWizard();
#endif
        if (dates.GetLength(0) != prices.GetLength(0))
        {
            return DExcelErrorMessage("Date array and stock price array have different sizes.");
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
        List<(DateTime date, double price)> sortedDatesAndPricesForVolCalculation = 
            sortedDatesAndPrices.Where(x => startDate <= x.date && x.date <= endDate).ToList();
        List<double> returns = new();
        for (int i = 0; i < sortedDatesAndPricesForVolCalculation.Count - 1; i++)
        {
            returns.Add(
                Math.Log(
                    sortedDatesAndPricesForVolCalculation[i].price / 
                    sortedDatesAndPricesForVolCalculation[i + 1].price));
        }

        double equallyWeightedVolatility = 0.0;
        double ewmaVolatility = 0.0;
        switch (weightingStyle.ToUpper())
        {
            case "EQUAL":
                equallyWeightedVolatility = mns.Statistics.StandardDeviation(returns) * Math.Sqrt(businessDaysPerYear);
                break;
            case "EXPONENTIAL":
            {
                List<double> squareReturns = returns.Select(x => x * x).ToList();
                double initialEwma = Math.Sqrt(squareReturns.Skip(Math.Max(0, squareReturns.Count - 24)).Sum());
                List<double> ewmaSeries = Enumerable.Repeat(0.0, squareReturns.Count - 1).Append(initialEwma).ToList();
                for (int i = ewmaSeries.Count - 1; i > 0; i--)
                {
                    ewmaSeries[i] = 
                        Math.Sqrt(
                            Math.Pow(ewmaSeries[i + 1], 2) * lambda + (1 - lambda) * squareReturns[i] * businessDaysPerYear);
                }
                ewmaVolatility = ewmaSeries[0];
                break;
            }
            default:
                return DExcelErrorMessage($"Invalid weighting style: {weightingStyle}");
        }

        return Math.Max(equallyWeightedVolatility, ewmaVolatility);
    }

    [ExcelFunction(
        Name = "d.Equity_CreateEuropeanOption",
        Description = "Creates an European equity option.",
        Category = "∂Excel: Equities")]
    public static string CreateEuropeanOption(
        [ExcelArgument(
            Name = "Handle",
            Description =
                "The 'handle' or name used to refer to the object in memory.\n" +
                "Each curve in a workbook must have a a unique handle.")]
        string handle,
        [ExcelArgument(Name = "S₀", Description = "Initial stock price.")]
        double spot,
        [ExcelArgument(Name = "K", Description = "Strike.")]
        double strike, 
        [ExcelArgument(Name = "r", Description = "Risk free rate (NACC). Only used for discounting.")]
        double riskFreeRate,
        [ExcelArgument(Name = "q", Description = "Dividend Yield (NACC).")]
        double dividendYield,
        [ExcelArgument(Name = "Vol", Description = "Volatility.")]
        double volatility,
        [ExcelArgument(Name = "Trade Date", Description = "The trade date.")]
        DateTime tradeDate,
        [ExcelArgument(Name = "Maturity Date", Description = "The maturity/exercise date of the option.")]
        DateTime maturityDate,
        [ExcelArgument(Name = "Calendar", Description = "The calendar for the option e.g., 'South Africa' or 'ZAR'.")]
        string calendar,
        [ExcelArgument(Name = "Day Count Convention", Description = "Day count convention e.g., 'Act360' or 'Act365'.")]
        string dayCountConvention,
        [ExcelArgument(Name = "Option Type", Description = "'Call'/'C' or 'Put'/'P'.")]
        string optionType)
    {
        if (!ParserUtils.TryParseQuantLibOptionType(
                optionType: optionType, 
                quantLibOptionType: out QL.Option.Type? quantLibOptionType,
                out string? optionTypeErrorMessage))
        {
            return optionTypeErrorMessage;
        }
        
        QL.PlainVanillaPayoff payoff = new((QL.Option.Type)quantLibOptionType, strike);
        QL.EuropeanExercise exercise = new(maturityDate.ToQuantLibDate()); 
        QL.VanillaOption vanillaOption = new(payoff, exercise);
        QL.QuoteHandle spotHandle = new(new QL.SimpleQuote(spot));
        
        ParserUtils.TryParseQuantLibDayCountConvention(dayCountConvention, out QL.DayCounter? dayCounter, out string errorMessage);

        QL.FlatForward interestRateCurve =
            new(tradeDate.ToQuantLibDate(),
                new QL.QuoteHandle(new QL.SimpleQuote(riskFreeRate)),
                dayCounter);
        
        QL.FlatForward dividendCurve =
            new(
                referenceDate: tradeDate.ToQuantLibDate(),
                forward: new QL.QuoteHandle(new QL.SimpleQuote(dividendYield)),
                dayCounter: dayCounter);

        (QL.Calendar? qlCalendar, string? qlCalendarErrorMessage) = DateUtils.ParseCalendars(calendar);
        if (qlCalendar == null)
        {
            return qlCalendarErrorMessage;
        }

        QL.BlackConstantVol constantVol = new(tradeDate.ToQuantLibDate(), qlCalendar, volatility, dayCounter);
        
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
        Name = "d.Equity_GetPrice",
        Description = "Get price of equity option.",
        Category = "∂Excel: Equities")]
    public static object GetPrice(string handle, bool verbose = false)
    {
        var dataObjectController = DataObjectController.Instance;
        object dataObject = dataObjectController.GetDataObject(handle);
        object output;

        if (dataObject.GetType() == typeof(QL.VanillaOption))
        {
            QL.VanillaOption vanillaOption = (QL.VanillaOption)dataObject;
            if (verbose)
            {
                output = new object[,]
                {
                    { "Price", vanillaOption.NPV() },
                    { "Delta", vanillaOption.delta() },
                    { "Vega", vanillaOption.vega() },
                };
            }
            else
            {
                output = vanillaOption.NPV();
            }
        }
        else if (dataObject.GetType() == typeof(QL.BasketOption))
        {
            QL.BasketOption basketOption = (QL.BasketOption)dataObject;
            if (verbose)
            {
                output = new object[,]
                {
                    { "Price", basketOption.NPV() },
                    { "Delta", basketOption.delta() },
                    { "Vega", basketOption.vega() },
                };
            }
            else
            {
                output = basketOption.NPV();
            }

        }
        else
        {
            return DExcelErrorMessage("Unknown option type.");
        }
        
        return output;
    }
}
