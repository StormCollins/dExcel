using System.IO;
using System.Reflection;
using ExcelDna.Integration;
using ExcelDna.Registration;
using dExcel.Dates;
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
        if (xllPath is null)
        {
            MessageBox.Show(@"∂Excel xll path not found.", @"∂Excel Error");
            throw new FileNotFoundException("∂Excel xll path not found.");
        }
        
        Assembly.LoadFrom(Path.Combine(xllPath, "dExcelWpf.dll"));
        Assembly.LoadFrom(Path.Combine(xllPath, "FuzzySharp.dll"));
        Assembly.LoadFrom(Path.Combine(xllPath, "MaterialDesignColors.dll"));
        Assembly.LoadFrom(Path.Combine(xllPath, "MaterialDesignThemes.Wpf.dll"));

        // Register the Standard Parameter Conversions (with the optional switch on how to treat references to empty cells)
        ParameterConversionConfiguration? paramConversionConfig =
            new ParameterConversionConfiguration()
                .AddParameterConversion(ParameterConversions.GetOptionalConversion(treatEmptyAsMissing: true))
                .AddParameterConversion((string s) => DateUtils.ParseDayCountConvention(s));
                // .AddParameterConversion((string s) => CommonUtils.ParseBusinessDayConvention(s));

        ExcelRegistration
            .GetExcelFunctions()
            .ProcessParamsRegistrations()
            .ProcessParameterConversions(paramConversionConfig)
            .RegisterFunctions();

        Task.Factory.StartNew(UsageStatsUtils.LogUsage);
    }
}
