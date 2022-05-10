namespace dExcel;

using System.Reflection;
using ExcelDna.Integration;

/// <summary>
/// A collection of utility functions to debug the dExcel application at runtime from Excel.
/// </summary>
public static class DebugUtils
{
    /// <summary>
    /// Gets the path of the dExcel xll.
    /// </summary>
    /// <returns>Path of the dExcel xll.</returns>
    [ExcelFunction(
        Name = "d.Debug_GetXllPath",
        Description = "Gets the path of the dExcel xll.",
        Category = "∂Excel: Debug")]
    public static string GetXllPath() => ExcelDnaUtil.XllPath;
}
