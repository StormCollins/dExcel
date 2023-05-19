namespace dExcel.CommonEnums;

/// <summary>
/// The common interpolation methods for curves e.g., interest rate and FX curves.
/// The name is delineated as "interpolation_method"_on_"rate/discount factor type" e.g., "Cubic_On_DiscountFactors"
/// means perform cubic interpolation on discount factors.
/// </summary>
public enum CurveInterpolationMethods
{
     // ReSharper disable InconsistentNaming
     Flat_On_ForwardRates,
     CubicSpline_On_DiscountFactors,
     Exponential_On_DiscountFactors,
     LogCubic_On_DiscountFactors,
     NaturalLogCubic_On_DiscountFactors,
     Cubic_On_ZeroRates,
     Linear_On_ZeroRates,
     NaturalCubic_On_ZeroRates,
}
