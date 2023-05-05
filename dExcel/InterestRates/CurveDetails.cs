using QL = QuantLib;

namespace dExcel.InterestRates;

/// <summary>
/// A class used to encapsulate the most common attributes of an interest rate curve.
/// </summary>
public class CurveDetails
{
    /// <summary>
    /// The term structure object which is used to calculate interpolated or node discount factors, zero rates, forward rates etc.
    /// </summary>
    public object? TermStructure { get; }
  
    /// <summary>
    /// The day count convention.
    /// </summary>
    public QL.DayCounter DayCountConvention { get; }
  
    /// <summary>
    /// The interpolation style of the curve used to interpolate discount factors.
    /// </summary>
    public string DiscountFactorInterpolation { get; } 
    
    /// <summary>
    /// The node dates of the discount factors. These are useful for plotting the curve for illustrative purposes.
    /// </summary>
    public List<DateTime>? DiscountFactorDates { get; }
   
    /// <summary>
    /// The discount factors at the node dates. These are useful for plotting the curve for illustrative purposes.
    /// </summary>
    public List<double>? DiscountFactors { get; }
    
    /// <summary>
    /// The instrument groups used to bootstrap the curve.
    /// </summary>
    public object[] InstrumentGroups { get; }

    /// <summary>
    /// The class constructor. 
    /// </summary>
    /// <param name="termStructure">The interest rate curve.</param>
    /// <param name="dayCountConvention">The day count convention.</param>
    /// <param name="interpolation">The interpolation style.</param>
    /// <param name="discountFactorDates">The discount factor dates.</param>
    /// <param name="discountFactors">The instrument groups. This is only populated if the curve was bootstrapped.
    /// </param>
    public CurveDetails(
        object? termStructure,
        QL.DayCounter dayCountConvention,
        string interpolation,
        IEnumerable<DateTime>? discountFactorDates,
        IEnumerable<double>? discountFactors,
        params object[] instrumentGroups)
    {
        this.TermStructure = termStructure;
        this.DayCountConvention = dayCountConvention;       
        this.DiscountFactorInterpolation = interpolation;
        this.DiscountFactorDates = discountFactorDates?.ToList();
        this.DiscountFactors = discountFactors?.ToList();
        this.InstrumentGroups = instrumentGroups;
    }
}
