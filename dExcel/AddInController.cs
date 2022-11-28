namespace dExcel;

using System.Data.SQLite;
using System.DirectoryServices;
using System.IO;
using System.Reflection;
using System.Security.Principal;
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

        if (NetworkUtils.GetConnectionStatus())
        {
            WindowsIdentity identity = WindowsIdentity.GetCurrent();
            string userName = identity.Name.Split('\\')[1];
            DirectoryEntry user = new($"LDAP://<SID={identity.User?.Value}>");
            user.RefreshCache(new[] { "givenName", "sn" });
            string? firstName = user.Properties["givenName"].Value?.ToString();
            string? surname = user.Properties["sn"].Value?.ToString();
            
            SQLiteConnection connection =
                new(
                    @"URI=file:\\\\ZAJNB010\Capital Markets 2\AQS Quants\dExcelTools\dExcelUsageStats\dexcel_usage_stats.sqlite");
            connection.Open();

            using SQLiteCommand writeCommand = new(connection);
            writeCommand.CommandText =
                $@"INSERT INTO users(username, firstname, surname, date_created, active)
               SELECT '{userName}', '{firstName}', '{surname}', DATETIME('NOW', 'localtime'), TRUE
               WHERE NOT EXISTS (SELECT * FROM users WHERE username='{userName}');";

            writeCommand.ExecuteNonQuery();
            string dexcelVersion = DebugUtils.GetAssemblyVersion();
            writeCommand.CommandText =
                $@"INSERT INTO dexcel_usage(username, version, date_logged)
               SELECT '{userName}', '{DebugUtils.GetAssemblyVersion()}', DATETIME('NOW', 'localtime')
               WHERE NOT EXISTS (SELECT * FROM dexcel_usage WHERE username='{userName}' AND version='{dexcelVersion}' AND DATE(date_logged)=DATE('NOW', 'localtime'));";

            writeCommand.ExecuteNonQuery();
        }
    }
}
