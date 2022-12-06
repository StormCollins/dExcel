namespace dExcel.InterestRates;

using QLNet;

public class CurveDetails
{
    public object? TermStructure { get; }
  
    public DayCounter DayCountConvention { get; }
   
    public IInterpolationFactory Interpolation { get; } 
    
    public List<Date>? Dates { get; }
    
    public List<double>? DiscountFactors { get; }

    public CurveDetails(
        object? termStructure,
        DayCounter dayCountConvention,
        IInterpolationFactory interpolation,
        IEnumerable<Date>? dates,
        IEnumerable<double>? discountFactors)
    {
        this.TermStructure = termStructure;
        this.DayCountConvention = dayCountConvention;       
        this.Interpolation = interpolation;
        this.Dates = dates?.ToList();
        this.DiscountFactors = discountFactors?.ToList();
    }
}
