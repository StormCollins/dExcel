namespace dExcel;

using System;
using System.Diagnostics;
using System.Windows.Navigation;
using System.Net.NetworkInformation;
using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

/// <summary>
/// Interaction logic for Dashboard.xaml
/// </summary>
public partial class Dashboard : Window
{
    /// <summary>
    /// A (possibly null) action that needs to be performed by Excel after the dashboard is closed.
    /// This is handled by the <see cref="RibbonController"/>.
    /// </summary>
    public string? DashBoardAction { get; set; } = null;

    private static Dashboard? _instance;
    private static bool _connectionStatus;
    private static bool _omicronStatus;
    

    public static Dashboard Instance
    {
        get
        {
            if (_instance == null)
            {
                _instance = new Dashboard();
            }
            return _instance;
        }
    }

    private Dashboard()
    {
        InitializeComponent();
        var xllPath = Path.GetDirectoryName(ExcelDnaUtil.XllPath);
        GitlabRepoLink.Source = new BitmapImage(new Uri(xllPath + @"\resources\icons\follow-link-small-green.ico"));
        InstallationPathLink.Source = new BitmapImage(new Uri(xllPath + @"\resources\icons\follow-link-small-green.ico"));
        JupyterHubLink.Source = new BitmapImage(new Uri(xllPath + @"\resources\icons\follow-link-small-green.ico"));
        this.InstalledDExcelVersion.Text = DebugUtils.GetAssemblyVersion();
        
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
        _instance = null;
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

    private void OpenInstaller_Click(object sender, RoutedEventArgs e)
    {
#if DEBUG
        var startupPath = Environment.ExpandEnvironmentVariables(@"%appdata%\Microsoft\Windows\Start Menu\Programs\Deloitte\dExcelInstaller.appref-ms");
        ProcessStartInfo startInfo = new()
        {
            UseShellExecute = true,
            FileName = "\"" + startupPath + "\""
        };

        Process process = new()
        {
            StartInfo = startInfo
        };

        process.Start();
#else
        Process.Start(@"C:\GitLab\dExcelTools\dExcel\dExcelInstaller\bin\Debug\net6.0-windows\dExcelInstaller.exe");
#endif
    }

    /// <summary>
    /// Opens the testing workbook: 'dexcel-testing.xlsm'.
    /// The logic is handled by <see cref="RibbonController.OpenDashboard"/> this just specifies the action.
    /// </summary>
    /// <param name="sender">The sender.</param>
    /// <param name="e">The routed event args.</param>
    private void OpenTestingWorkbook_Click(object sender, RoutedEventArgs e)
    {
        this.DashBoardAction = "OpenTestingWorkbook";
        this.Close();
    }
}
