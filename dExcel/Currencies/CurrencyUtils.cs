namespace dExcel.Currencies;

using System.Reflection;
using QLNet;

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
    public static Currency? ParseCurrency(string currencyToParse)
    {
        Assembly? qlNet = 
            AppDomain.CurrentDomain.GetAssemblies().SingleOrDefault(assembly => assembly.GetName().Name == "QLNet");
        
        Type? type = qlNet?.GetType($"QLNet.{currencyToParse.ToUpper()}Currency");
        return type is not null ? (Currency?) Activator.CreateInstance(type) : null;
    }
}
