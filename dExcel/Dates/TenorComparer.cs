using Omicron;

namespace dExcel.Dates;

/// <summary>
/// A class implementing a comparer for Omicron tenors.
/// </summary>
public class TenorComparer : Comparer<Tenor>
{
    /// <summary>
    /// Compares two tenors.
    /// </summary>
    /// <param name="x">First tenor.</param>
    /// <param name="y">Second tenor.</param>
    /// <exception cref="ArgumentOutOfRangeException">Thrown if an unknown tenor is used.</exception>
    /// <returns>0 if x == y, -1 if x &lt; y, and 1 if x &gt; y.</returns>
    public override int Compare(Tenor? x, Tenor? y)
    {
        if (x == null && y == null) return 0;
       
        if (x == null) return -1;
        
        if (y == null) return 1;
        
        if (x.Unit == y.Unit)
        {
            return x.Amount.CompareTo(y.Amount);
        }

        int xAmount = x.Unit switch
        {
            TenorUnit.Day => x.Amount,
            TenorUnit.Week => x.Amount * 7,
            TenorUnit.Month => x.Amount * 30,
            TenorUnit.Year => x.Amount * 365,
            _ => throw new ArgumentOutOfRangeException(),
        };

        int yAmount = y.Unit switch
        {
            TenorUnit.Day => y.Amount,
            TenorUnit.Week => y.Amount * 7,
            TenorUnit.Month => y.Amount * 30,
            TenorUnit.Year => y.Amount * 365,
            _ => throw new ArgumentOutOfRangeException(),
        };
        
        return xAmount.CompareTo(yAmount);   
    }
}
