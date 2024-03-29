﻿using QL = QuantLib;
using System.Reflection;

namespace dExcel.FX;

/// <summary>
/// A set of utility functions for dealing with currencies.
/// </summary>
public static class CurrencyUtils
{
    /// <summary>
    /// Parses a string as a QuantLib currency.
    /// </summary>
    /// <param name="currencyToParse">Currency to parse.</param>
    /// <returns>QuantLib currency.</returns>
    public static QL.Currency? ParseCurrency(string currencyToParse)
    {
        Assembly? quantLib = 
            AppDomain.CurrentDomain.GetAssemblies().SingleOrDefault(assembly => assembly.GetName().Name == "NQuantLib");
        
        Type? type = quantLib?.GetType($"QuantLib.{currencyToParse.ToUpper()}Currency");
        return type is not null ? (QL.Currency?) Activator.CreateInstance(type) : null;
    }
}
