using System.Data.SQLite;
using System.DirectoryServices;
using System.IO;
using System.Reflection;
using System.Security.Principal;
using ExcelDna.Integration;
using ExcelDna.Registration;
using dExcel.Utilities;

namespace dExcel;

public class AddInController : IExcelAddIn
{
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

        // Register the Standard Parameter Conversions (with the optional switch on how to treat references to empty cells)
        ParameterConversionConfiguration? paramConversionConfig =
            new ParameterConversionConfiguration()
                .AddParameterConversion(ParameterConversions.GetOptionalConversion(treatEmptyAsMissing: true));

        ExcelRegistration
            .GetExcelFunctions()
            .ProcessParamsRegistrations()
            .ProcessParameterConversions(paramConversionConfig)
            .RegisterFunctions();

        Task.Factory.StartNew(UsageStatsUtils.LogUsage);
    }
}
