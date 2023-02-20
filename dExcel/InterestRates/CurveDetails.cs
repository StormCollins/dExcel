namespace dExcel.InterestRates;

using QLNet;

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
    public DayCounter DayCountConvention { get; }
  
    /// <summary>
    /// The interpolation style of the curve used to interpolate discount factors.
    /// </summary>
    public string DiscountFactorInterpolation { get; } 
    
    /// <summary>
    /// The node dates of the discount factors.
    /// </summary>
    public List<Date>? DiscountFactorDates { get; }
   
    /// <summary>
    /// The discount factors at the node dates.
    /// </summary>
    public List<double>? DiscountFactors { get; }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="termStructure"></param>
    /// <param name="dayCountConvention"></param>
    /// <param name="interpolation"></param>
    /// <param name="discountFactorDates"></param>
    /// <param name="discountFactors"></param>
    public CurveDetails(
        object? termStructure,
        DayCounter dayCountConvention,
        string interpolation,
        IEnumerable<Date>? discountFactorDates,
        IEnumerable<double>? discountFactors)
    {
        this.TermStructure = termStructure;
        this.DayCountConvention = dayCountConvention;       
        this.DiscountFactorInterpolation = interpolation;
        this.DiscountFactorDates = discountFactorDates?.ToList();
        this.DiscountFactors = discountFactors?.ToList();
    }
}
