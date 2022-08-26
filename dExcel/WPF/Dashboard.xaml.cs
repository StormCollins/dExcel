namespace dExcel;

using System;
using System.Diagnostics;
using System.Windows.Navigation;
using System.Net.NetworkInformation;
using System.IO;
using System.Text.RegularExpressions;
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
            this.LatestDExcelVersion.Text = GetAllAvailableRemoteDExcelReleases()?.Max();
        }
        else
        {
            _connectionStatus = false;
            ConnectionStatus.Source = new BitmapImage(new Uri(xllPath + "/resources/icons/connection-status-amber.ico"));
            this.LatestDExcelVersion.Text = GetAllAvailableLocalDExcelReleases().Max();
        }

        if (NetworkUtils.GetOmicronStatus())
        {
            _omicronStatus = true;
            this.OmicronStatus.Source = new BitmapImage(new Uri(xllPath + "/resources/icons/database-connected-large-green.ico"));
        }
        else
        {
            _omicronStatus = false;
            this.OmicronStatus.Source = new BitmapImage(new Uri(xllPath + "/resources/icons/database-not-connected-large-amber.ico"));
        }

        NetworkChange.NetworkAddressChanged += ConnectionStatusChangedCallback;
    }
    /// <summary>
    /// The location of all the add-in versions on the shared drive.
    /// </summary>
    private const string SharedDriveReleasesPath = @"\\ZAJNB010\Capital Markets 2\AQS Quants\dExcelTools\Releases";

    /// <summary>
    /// The location of the workbook which invokes a VBA based installer on start up.
    /// </summary>
    private const string SharedDriveInstallationWorkbookPath = @"\\ZAJNB010\Capital Markets 2\AQS Quants\dExcelTools\Installer\dExcelInstaller.xlsm";

    /// <summary>
    /// The location of the all the add-in versions on the local machine.
    /// </summary>
    private const string LocalReleasesPath = @"C:\GitLab\dExcelTools\Releases\";
    
    /// <summary>
    /// The location of the currently installed version of the add-in on the machine.
    /// </summary>
    private const string LocalCurrentReleasePath = @"C:\GitLab\dExcelTools\Releases\Current\";

    /// <summary>
    /// Gets all versions of the ∂Excel add-in already copied to the user's local machine.
    /// </summary>
    /// <returns>All versions of the ∂Excel add-in available locally.</returns>
    private static IEnumerable<string?> GetAllAvailableLocalDExcelReleases()
        => Directory
            .GetDirectories(LocalReleasesPath)
            .Where(x => Regex.IsMatch(x, @"\d+(.\d+)"))
            .Select(Path.GetFileName)
            .Reverse();

    /// <summary>
    /// Gets all versions of the ∂Excel add-in at the specified remote source location e.g. the shared drive or GitLab.
    /// </summary>
    /// <returns>All versions of the ∂Excel add-in available remotely.</returns>
    private IEnumerable<string?>? GetAllAvailableRemoteDExcelReleases()
    {
        return Directory
            .GetFiles(SharedDriveReleasesPath)
            .Where(x => Regex.IsMatch(x, @"\d+(.\d+)(?=\.zip)"))
            .Select(Path.GetFileNameWithoutExtension)
            .Reverse();
    }

    private void Dashboard_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        _instance = null;
    }

    private void CloseDashboard(object sender, RoutedEventArgs e)
    {
        this.Close();
    }

    private void ConnectionStatusChangedCallback(object sender, EventArgs e)
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
                    this.LatestDExcelVersion.Text = GetAllAvailableRemoteDExcelReleases()?.Max();
                });
            }
            else
            {
                Dispatcher.Invoke(() =>
                {
                    ConnectionStatus.Source = new BitmapImage(new Uri(xllPath + "/resources/icons/connection-status-amber.ico"));
                    this.LatestDExcelVersion.Text = GetAllAvailableLocalDExcelReleases().Max();
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
        Process.Start(@"C:\GitLab\dExcelTools\dExcel\dExcelInstaller\bin\Debug\net6.0-windows\dExcelInstaller.exe");
#else
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
