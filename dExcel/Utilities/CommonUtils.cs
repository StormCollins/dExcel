namespace dExcel;

using ExcelDna.Integration;

public static class CommonUtils
{
    /// <summary>
    /// Returns a message stating that the user is currently in the function wizard dialog in Excel rather than pre-computing values.
    /// Used when pre-computing values is expensive and the user is entering inputs one at a time.
    /// </summary>
    /// <returns>'In a function wizard.'</returns>
    public static string InFunctionWizard() => ExcelDnaUtil.IsInFunctionWizard() ? "In function wizard." : "";

    /// <summary>
    /// The prefix used by error messages from ∂Excel.
    /// </summary>
    public const string DExcelErrorPrefix = "#∂Excel Error:";
}
