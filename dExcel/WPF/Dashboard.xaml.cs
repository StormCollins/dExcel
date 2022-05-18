namespace dExcel;

using System;
using System.Diagnostics;
using System.Windows.Navigation;
using System.Net.NetworkInformation;
using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;
using ExcelDna.Integration;
using MaterialDesignThemes.Wpf;

/// <summary>
/// Interaction logic for Dashboard.xaml
/// </summary>
public partial class Dashboard : Window
{
    private static Dashboard? instance = null;
    private static bool _connectionStatus;
    private static bool _omicronStatus;

    public static Dashboard Instance
    {
        get
        {
            if (instance == null)
            {
                instance = new Dashboard();
            }
            return instance;
        }
    }

    private Dashboard()
    {
        var xllPath = Path.GetDirectoryName(ExcelDnaUtil.XllPath);
        InitializeComponent();
        ShadowAssist.SetShadowDepth(this, ShadowDepth.Depth0);
        this.Icon = dExcelIcon.Source = new BitmapImage(new Uri(xllPath + @"\resources\icons\dXL-logo-extra-small.ico")); 
        dExcelIcon.Source = new BitmapImage(new Uri(xllPath + @"\resources\icons\dExcel.ico"));
        gitlabRepoLink.Source = new BitmapImage(new Uri(xllPath + @"\resources\icons\follow-link-small-green.ico"));
        installationPathLink.Source = new BitmapImage(new Uri(xllPath + @"\resources\icons\follow-link-small-green.ico"));
        Closing += Dashboard_Closing;
        if (NetworkUtils.GetConnectionStatus())
        {
            _connectionStatus = true;
            ConnectionStatus.Source = new BitmapImage(new Uri(xllPath + "/resources/icons/connection-status-green.ico"));
        }
        else
        {
            _connectionStatus = false;
            ConnectionStatus.Source = new BitmapImage(new Uri(xllPath + "/resources/icons/connection-status-amber.ico"));
        }

        if (NetworkUtils.GetOmicronStatus())
        {
            _omicronStatus = true;
            OmicronStatus.Source = new BitmapImage(new Uri(xllPath + "/resources/icons/database-connected-large-green.ico"));
        }
        else
        {
            _omicronStatus = false;
            OmicronStatus.Source = new BitmapImage(new Uri(xllPath + "/resources/icons/database-not-connected-large-amber.ico"));
        }

        NetworkChange.NetworkAddressChanged += ConnectionStatusChangedCallback;
    }

    private void Dashboard_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        instance = null;
    }

    private void CloseDashboard(object sender, RoutedEventArgs e)
    {
        this.Close();
    }

    public void ConnectionStatusChangedCallback(object sender, EventArgs e)
    {
        var xllPath = Path.GetDirectoryName(ExcelDnaUtil.XllPath);
        if (NetworkUtils.GetConnectionStatus() != _connectionStatus)
        {
            _connectionStatus = !_connectionStatus;
            if (NetworkUtils.GetConnectionStatus())
            {
                Dispatcher.Invoke(() =>
                {
                    ConnectionStatus.Source = new BitmapImage(new Uri(xllPath + "/resources/icons/connection-status-green.ico"));
                });
            }
            else
            {
                Dispatcher.Invoke(() =>
                {
                    ConnectionStatus.Source = new BitmapImage(new Uri(xllPath + "/resources/icons/connection-status-amber.ico"));
                });
            }
        }

        if (NetworkUtils.GetOmicronStatus() != _omicronStatus)
        {
            _omicronStatus = !_omicronStatus;
            if (NetworkUtils.GetOmicronStatus())
            {
                Dispatcher.Invoke(() =>
                {
                    OmicronStatus.Source = new BitmapImage(new Uri(xllPath + "/resources/icons/database-connected-large-green.ico"));
                });
            }
            else
            {
                Dispatcher.Invoke(() =>
                {
                    OmicronStatus.Source = new BitmapImage(new Uri(xllPath + @"\resources\icons\database-not-connected-large-amber.ico"));
                });
            }
        }
    }

    private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
    {
        Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true });
        e.Handled = true;
    }
}
