namespace dExcel;

using System.IO;
using System.Reflection;
using ExcelDna.Integration;
using ExcelDna.Registration;
using Excel = Microsoft.Office.Interop.Excel;

public class AddInController : IExcelAddIn
{
    Excel.Application xlapp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
    
    public void AutoClose()
    {
        
    }

    public void AutoOpen()
    {
        string? xllPath = Path.GetDirectoryName(ExcelDnaUtil.XllPath);
        Assembly.LoadFrom(Path.Combine(xllPath, "dExcelWpf.dll"));
        Assembly.LoadFrom(Path.Combine(xllPath, "FuzzySharp.dll"));
        Assembly.LoadFrom(Path.Combine(xllPath, "MaterialDesignColors.dll"));
        Assembly.LoadFrom(Path.Combine(xllPath, "MaterialDesignThemes.Wpf.dll"));
        Assembly.LoadFrom(Path.Combine(xllPath, "QLNet.dll"));

        // Register the Standard Parameter Conversions (with the optional switch on how to treat references to empty cells)
        var paramConversionConfig =
            new ParameterConversionConfiguration()
                .AddParameterConversion(ParameterConversions.GetOptionalConversion(treatEmptyAsMissing: true));

        ExcelRegistration
            .GetExcelFunctions()
            .ProcessParamsRegistrations()
            .ProcessParameterConversions(paramConversionConfig)
            .RegisterFunctions();
    }
}
