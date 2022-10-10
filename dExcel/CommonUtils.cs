namespace dExcel;

using ExcelDna.Integration;
using QLNet;

public static class CommonUtils
{
    /// <summary>
    /// Returns a message stating that the user is currently in the function wizard dialog in Excel rather than pre-computing values.
    /// Used when pre-computing values is expensive and the user is entering inputs one at a time.
    /// </summary>
    /// <returns>'In a function wizard.'</returns>
    public static string InFunctionWizard() => ExcelDnaUtil.IsInFunctionWizard() ? "In function wizard." : "";

    /// <summary>
    /// Users can get available <see cref="GetAvailableBusinessDayConventions"/>
    /// </summary>
    /// <param name="businessDayConventionToParse"></param>
    /// <returns></returns>
    public static BusinessDayConvention? ParseBusinessDayConvention(string businessDayConventionToParse)
    {
        BusinessDayConvention? businessDayConvention = businessDayConventionToParse.ToUpper() switch
        {
            "FOLLOWING" or "FOL" => BusinessDayConvention.Following,
            "MODIFIEDFOLLOWING" or "MODFOL" => BusinessDayConvention.ModifiedFollowing,
            "MODIFIEDPRECEDING" or "MODPREC" => BusinessDayConvention.ModifiedPreceding,
            "PRECEDING" or "PREC" => BusinessDayConvention.Preceding,
            _ => null,
        };

        return businessDayConvention;
    }

}
