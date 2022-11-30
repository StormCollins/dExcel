using dExcel.InterestRates;

namespace dExcelTests;

using dExcel;
using NUnit.Framework;
using QLNet;

[TestFixture]
public class HullWhiteTests
{
    [TestCase]
    public void Calibration()
    { 
        object[,] curveParameters =
        {
            { "Parameter", "Value" },
            { "DayCountConvention", "Actual365" },
            { "Interpolation", "LogLinear" },
        }; 
        
        var dates = new object[,]
        {
            { new DateTime(2022,  05, 25).ToOADate() },
            { new DateTime(2022,  05, 26).ToOADate() },
            { new DateTime(2022,  06, 06).ToOADate() },
            { new DateTime(2022,  06, 29).ToOADate() },
            { new DateTime(2022,  07, 31).ToOADate() },
            { new DateTime(2022,  08, 29).ToOADate() },
            { new DateTime(2022,  09, 26).ToOADate() },
            { new DateTime(2022,  10, 23).ToOADate() },
            { new DateTime(2022,  11, 22).ToOADate() },
            { new DateTime(2022,  12, 25).ToOADate() },
            { new DateTime(2023,  01, 23).ToOADate() },
            { new DateTime(2023,  02, 22).ToOADate() },
            { new DateTime(2023,  03, 27).ToOADate() },
            { new DateTime(2023,  04, 25).ToOADate() },
            { new DateTime(2023,  05, 25).ToOADate() },
            { new DateTime(2023,  08, 24).ToOADate() },
            { new DateTime(2023,  11, 22).ToOADate() },
            { new DateTime(2024,  02, 22).ToOADate() },
            { new DateTime(2024,  05, 26).ToOADate() },
            { new DateTime(2025,  05, 29).ToOADate() },
            { new DateTime(2026,  05, 31).ToOADate() },
            { new DateTime(2027,  05, 30).ToOADate() },
            { new DateTime(2028,  05, 28).ToOADate() },
            { new DateTime(2029,  05, 29).ToOADate() },
            { new DateTime(2030,  05, 29).ToOADate() },
            { new DateTime(2031,  05, 29).ToOADate() },
            { new DateTime(2032,  05, 30).ToOADate() },
            { new DateTime(2034,  05, 29).ToOADate() },
            { new DateTime(2037,  05, 31).ToOADate() },
            { new DateTime(2042,  05, 29).ToOADate() },
            { new DateTime(2047,  05, 29).ToOADate() },
            { new DateTime(2052,  05, 28).ToOADate() },
            { new DateTime(2062,  05, 29).ToOADate() },
        };
        
        var discountFactors = new object[,] 
        {
            { 1.000000 },
            { 0.999998 },
            { 0.999976 },
            { 0.999923 },
            { 0.999796 },
            { 0.999658 },
            { 0.999585 },
            { 0.999421 },
            { 0.999249 },
            { 0.999215 },
            { 0.999062 },
            { 0.998858 },
            { 0.998781 },
            { 0.998564 },
            { 0.998309 },
            { 0.997485 },
            { 0.996260 },
            { 0.994688 },
            { 0.992613 },
            { 0.980440 },
            { 0.965201 },
            { 0.948184 },
            { 0.930458 },
            { 0.912260 },
            { 0.894270 },
            { 0.876225 },
            { 0.858578 },
            { 0.822754 },
            { 0.773045 },
            { 0.696414 },
            { 0.632344 },
            { 0.579070 },
            { 0.499731 },
        };

        string curveHandle = Curve.Create("DiscountCurve", curveParameters, dates, discountFactors);
        
        var swaptionMaturities = new object[,] 
        { 
            { 7 },
            { 5 },
            { 4 },
            { 3 },
            { 2 }, 
            { 1 },
        };
        
        var swapLengths = new object[,]
        {
            {1},
            {3},
            {4},
            {5},
            {6},
            {7}
        };
        
        var swaptionVols = new object[,]
        { 
            {0.362},
            {0.380},
            {0.389},
            {0.402}, 
            {0.421},
            {0.457},
        };
        
        // var parameters = dExcel.HullWhite.Calibrate(curveHandle, swaptionMaturities, swapLengths, swaptionVols);
        
    }
}
