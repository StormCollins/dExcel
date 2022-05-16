namespace dExcelInstaller;

using System;
using System.Net;

public static class NetworkUtils
{
    /// <summary>
    /// Gets a boolean value indicating whether the user is currently connected to the Deloitte network.
    /// </summary>
    /// <returns>True if connected to network otherwise false.</returns>
    public static bool GetConnectionStatus()
    {
        try
        {
            IPHostEntry host = Dns.GetHostEntry(@"gitlab.fsa-aks.deloitte.co.za");
            return true;
        }
        catch (Exception)
        {
            return false;
        }
    }
}
