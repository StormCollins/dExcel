using QL = QuantLib;
using System.Reflection;

namespace dExcel.Currencies;

/// <summary>
/// A set of utility functions for dealing with currencies.
/// </summary>
public static class CurrencyUtils
{
    /// <summary>
    /// Parses a string as a QLNet currency.
    /// </summary>
    /// <param name="currencyToParse">Currency to parse.</param>
    /// <returns>QLNet currency.</returns>
    public static QL.Currency? ParseCurrency(string currencyToParse)
    {
        Assembly? quantlib = 
            AppDomain.CurrentDomain.GetAssemblies().SingleOrDefault(assembly => assembly.GetName().Name == "NQuantLib");
        
        Type? type = quantlib?.GetType($"QuantLib.{currencyToParse.ToUpper()}Currency");
        return type is not null ? (QL.Currency?) Activator.CreateInstance(type) : null;
    }
}
