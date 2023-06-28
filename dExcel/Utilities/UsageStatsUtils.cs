namespace dExcel.Utilities;

using System.Data.SQLite;
using System.DirectoryServices;
using System.Security.Principal;
using ExcelUtils;

internal static class UsageStatsUtils
{

    public static void LogUsage()
    {
        if (NetworkUtils.GetVpnConnectionStatus())
        {
            WindowsIdentity identity = WindowsIdentity.GetCurrent();
            string userName = identity.Name.Split('\\')[1];
            DirectoryEntry user = new($"LDAP://<SID={identity.User?.Value}>");
            user.RefreshCache(new[] { "givenName", "sn" });
            string? firstName = user.Properties["givenName"].Value?.ToString();
            string? surname = user.Properties["sn"].Value?.ToString();
                
            try
            {
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
            catch (Exception e)
            {
                MessageBox.Show(e.Message, @"∂Excel Error");
            }
        }
    }
}
