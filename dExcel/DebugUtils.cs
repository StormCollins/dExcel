namespace dExcel;

using System.Reflection;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

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

    /// <summary>
    /// Gets a list of all assemblies loaded into the current domain.
    /// </summary>
    /// <returns>A list of all assemblies loaded into the current domain.</returns>
    [ExcelFunction(
        Name = "d.Debug_GetAssemblies",
        Description = "Gets a List of all assemblies loaded into the current domain.",
        Category = "∂Excel: Debug")]
    public static object[,] GetAssemblies()
    {
        var assemblyNames = new List<string>();
        foreach (var assembly in AppDomain.CurrentDomain.GetAssemblies())
        {
            if (assembly.FullName != null)
            {
                assemblyNames.Add(assembly.FullName);
            }
        }

        assemblyNames.Sort();

        var output = new object[assemblyNames.Count, 1];
        for (int i = 0; i < assemblyNames.Count; i++)
        {
            output[i, 0] = assemblyNames[i];
        }
        return output;
    }

    [ExcelFunction(
        Name = "d.Debug_GetAddInPath",
        Description = "Gets a List of all assemblies loaded into the current domain.",
        Category = "∂Excel: Debug")]
    public static string GetAddInPath()
    {
        var xlApp = (Excel.Application)ExcelDnaUtil.Application;
        foreach (Excel.AddIn addIn in xlApp.AddIns)
        {
            if (addIn.FullName.Contains("dExcel"))
            {
                return addIn.FullName;
            }
        }
        return "";
    }

    [ExcelFunction(
        Name = "d.Debug_GetAssemblyVersion",
        Description = "Gets a List of all assemblies loaded into the current domain.",
        Category = "∂Excel: Debug")]
    public static string GetAssemblyVersion()
    {
        var dExcelAssembly = Assembly.GetAssembly(typeof(DebugUtils))?.GetName().Version;
        return dExcelAssembly == null ? "Failed to get version." : $"{dExcelAssembly.Major}.{dExcelAssembly.Minor}.{dExcelAssembly.Build}";
    }
}
