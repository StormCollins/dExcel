namespace dExcel.Utilities;

using System;
using System.Net;

public static class NetworkUtils
{
    /// <summary>
    /// Gets a boolean value indicating whether the user is currently connected to the Deloitte network.
    /// </summary>
    /// <returns>True if connected to the Deloitte network otherwise false.</returns>
    public static bool GetVpnConnectionStatus()
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

    /// <summary>
    /// Gets a boolean value indicating whether the user can currently connect to Omicron.
    /// </summary>
    /// <returns>True if the user can connect to Omicron otherwise false.</returns>
    public static bool GetOmicronStatus()
    {
        try
        {
            IPHostEntry host = Dns.GetHostEntry(@"omicron.fsa-aks.deloitte.co.za");
            return true;
        }
        catch (Exception)
        {
            return false;
        }
    }
}
