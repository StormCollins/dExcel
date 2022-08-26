namespace dExcelInstaller;

using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Security.Principal;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    /// <summary>
    /// The directory containing dExcel releases, local or remote, to monitor for new releases.
    /// </summary>
    private readonly FileSystemWatcher _releasesDirectoryWatcher = new();
    
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
    /// The dExcel dll name.
    /// </summary>
    private const string Dll = "dExcel.dll";
    
    /// <summary>
    /// Indicates if the user is connected to the VPN.
    /// </summary>
    private static bool _vpnConnectionStatus;

    /// <summary>
    /// The logging window housed in the installer.
    /// </summary>
    private readonly Logger _logger;

    /// <summary>
    /// The main window for the installer.
    /// </summary>
    public MainWindow()
    {
        InitializeComponent();
        this._logger = new Logger(LogWindow);
        var installerVersion = Assembly.GetEntryAssembly()?.GetName().Version;
        InstallerVersion.Output = $"{installerVersion?.Major}.{installerVersion?.Minor}";

        var currentDExcelVersion = GetInstalledDExcelVersion();
        CurrentDExcelVersion.Output = currentDExcelVersion;

        if (string.Compare(
                strA: currentDExcelVersion,
                strB: "Not Installed",
                comparisonType: StringComparison.InvariantCultureIgnoreCase) == 0)
        {
            this.Uninstall.IsEnabled = false;
        }

        AdminRights.Text = IsAdministrator().ToString();
        if (!IsAdministrator())
        {
            this._logger.WarningText =
                "The current user is not an admin on this machine. It may limit installation functionality.";
        }

        if (NetworkUtils.GetConnectionStatus())
        {
            _vpnConnectionStatus = true;
            this.ConnectionStatus.Source =
                new BitmapImage(
                    new Uri(@"pack://application:,,,/resources/icons/connection-status-green.ico",
                        UriKind.Absolute));
            this._releasesDirectoryWatcher.Path = SharedDriveReleasesPath;
            this.DockPanelConnectionStatus.ToolTip = "You are connected to the VPN.";
            this._logger.OkayText =
                $"Checking for latest versions of ∂Excel on the selected remote source: **{DExcelRemoteSource.Text}**";
            this._logger.WarningText = $"Installation path set to: [[{SharedDriveReleasesPath}]]";
            try
            {
                this.AvailableDExcelReleases.ItemsSource = GetAllAvailableRemoteDExcelReleases();
            }
            catch (Exception e)
            {
                this._logger.ErrorText = e.Message;
            }
        }
        else
        {
            _vpnConnectionStatus = false;
            this.ConnectionStatus.Source =
                new BitmapImage(
                    new Uri(
                        uriString: @"pack://application:,,,/resources/icons/connection-status-amber.ico",
                        uriKind: UriKind.Absolute));
            this._releasesDirectoryWatcher.Path = LocalReleasesPath;
            this._logger.WarningText = "User not connected to the VPN.";
            this._logger.WarningText =
                "The VPN is required to check for the latest versions of the ∂Excel add-in on the selected remote " +
                $"source: **{this.DExcelRemoteSource.Text}**";
            this._logger.WarningText = $"Installation path set to: [[{LocalReleasesPath}]]";
            AvailableDExcelReleases.ItemsSource = GetAllAvailableLocalDExcelReleases();
            this.DockPanelConnectionStatus.ToolTip = "You are not connected to the VPN.";
        }

        AvailableDExcelReleases.SelectedIndex = 0;
        NetworkChange.NetworkAddressChanged += ConnectionStatusChangedCallback!;
        _releasesDirectoryWatcher.NotifyFilter = 
            NotifyFilters.Attributes |
            NotifyFilters.CreationTime |
            NotifyFilters.DirectoryName |
            NotifyFilters.FileName | 
            NotifyFilters.LastAccess |
            NotifyFilters.LastWrite |
            NotifyFilters.Security |
            NotifyFilters.Size;
        
        _releasesDirectoryWatcher.Changed += ReleasesesFolderChanged;
        _releasesDirectoryWatcher.Deleted += ReleasesesFolderChanged;
        _releasesDirectoryWatcher.Filter = "*.*";
        _releasesDirectoryWatcher.IncludeSubdirectories = true;
        _releasesDirectoryWatcher.EnableRaisingEvents = true; 
    }

    private void ReleasesesFolderChanged(object sender, FileSystemEventArgs e)
    {
        Dispatcher.Invoke(() =>
        {
            if (_releasesDirectoryWatcher.Path == SharedDriveReleasesPath)
            {
                this.AvailableDExcelReleases.ItemsSource = GetAllAvailableRemoteDExcelReleases();
            }
            else
            {
                this.AvailableDExcelReleases.ItemsSource = GetAllAvailableLocalDExcelReleases();
            }
            this.AvailableDExcelReleases.SelectedIndex = 0;
        });
    }

    /// <summary>
    /// Checks if the current user is an administrator on the current machine.
    /// </summary>
    /// <returns>Returns true if the user is an administrator otherwise false.</returns>
    private static bool IsAdministrator()
    {
        using var identity = WindowsIdentity.GetCurrent();
        var principal = new WindowsPrincipal(identity);
        return principal.IsInRole(WindowsBuiltInRole.Administrator);
    }

    /// <summary>
    /// Callback triggered by the either the VPN connection status or Omicron connection status changing.
    /// </summary>
    /// <param name="sender">The sender.</param>
    /// <param name="e">The EventArgs.</param>
    private void ConnectionStatusChangedCallback(object sender, EventArgs e)
    {
        if (NetworkUtils.GetConnectionStatus() != _vpnConnectionStatus)
        {
            _vpnConnectionStatus = !_vpnConnectionStatus;
            if (NetworkUtils.GetConnectionStatus())
            {
                Dispatcher.Invoke(() =>
                {
                    this.ConnectionStatus.Source =
                        new BitmapImage(
                            new Uri(
                                uriString: @"pack://application:,,,/resources/icons/connection-status-green.ico",
                                uriKind: UriKind.Absolute));
                    this._releasesDirectoryWatcher.Path = SharedDriveReleasesPath;
                    this.DockPanelConnectionStatus.ToolTip = "You are connected to the VPN.";
                    this._logger.OkayText =
                        "Checking for latest versions of ∂Excel on the selected remote source: " +
                        $"**{this.DExcelRemoteSource.Text}**";
                    this._logger.OkayText = $"Installation path set to: [[{SharedDriveReleasesPath}]]";
                    this.AvailableDExcelReleases.ItemsSource = GetAllAvailableRemoteDExcelReleases();
                });
            }
            else
            {
                Dispatcher.Invoke(() =>
                {
                    this.ConnectionStatus.Source =
                        new BitmapImage(
                            new Uri(
                                uriString: @"pack://application:,,,/resources/icons/connection-status-amber.ico",
                                uriKind: UriKind.Absolute));
                    this._releasesDirectoryWatcher.Path = LocalReleasesPath;
                    this.DockPanelConnectionStatus.ToolTip = "You are not connected to the VPN.";
                    this._logger.WarningText = "User not connected to VPN.";
                    this._logger.WarningText = 
                        "The VPN is required to check for latest versions of the ∂Excel add-in on the selected " +
                        $"remote source: **{DExcelRemoteSource.Text}**";
                    this._logger.WarningText =
                        "Only locally available versions of the ∂Excel add-in can be installed.";
                    this._logger.WarningText = $"Installation path set to: [[{LocalReleasesPath}]]";
                    this.AvailableDExcelReleases.ItemsSource = GetAllAvailableLocalDExcelReleases();
                });
            }

            Dispatcher.Invoke(() =>
            {
                this.AvailableDExcelReleases.SelectedIndex = 0;
            });
        }
    }

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
        if (string.Compare(DExcelRemoteSource.Text, "Shared Drive", StringComparison.OrdinalIgnoreCase) == 0)
        {
            return Directory
                .GetFiles(SharedDriveReleasesPath)
                .Where(x => Regex.IsMatch(x, @"\d+(.\d+)(?=\.zip)"))
                .Select(Path.GetFileNameWithoutExtension)
                .Reverse();
        }

        return null;
    }
    
    /// <summary>
    /// Installation process triggered by clicking the install button.
    /// </summary>
    /// <param name="sender">The sender.</param>
    /// <param name="e">The routed events.</param>
    private void Install_Click(object sender, RoutedEventArgs e)
    {
        new Thread(InstallAddIn).Start();
    }

    /// <summary>
    /// Gets the version of dExcel currently installed in Excel.
    /// </summary>
    /// <returns>The dExcel version if available as "Major version number.Minor version number" if installed otherwise
    /// "Not Installed".</returns>
    private static string GetInstalledDExcelVersion()
    {
        if (File.Exists(LocalCurrentReleasePath + @"\" + Dll))
        {
            var currentDExcelVersion = AssemblyName.GetAssemblyName(LocalCurrentReleasePath + @"\" + Dll).Version;
            return $"{currentDExcelVersion?.Major}.{currentDExcelVersion?.Minor}";
        }

        return "Not Installed";
    }
    
    TaskCompletionSource<bool> tcs = new();
    
    private async void IgnoreExcelIsOpenWarning_OnClick(object sender, RoutedEventArgs e)
    {
        // tcs = new TaskCompletionSource<bool>();
        tcs.SetResult(true);
    }
    

    private void StopInstallationAndDontCloseExcel_OnClick(object sender, RoutedEventArgs e)
    {
        tcs.SetResult(false);
    }
    
    /// <summary>
    /// Installs the specified version of the dExcel AddIn to Excel.
    /// </summary>
    private async void InstallAddIn()
    {
        // Check with user if all unsaved instances of Excel can be terminated.
        if (Process.GetProcessesByName("Excel").Any())
        {
            await Dispatcher.Invoke(async () =>
            {
                this.ExcelIsOpenWarningDialog.IsOpen = true;
                await tcs.Task;
            });
            
            if (tcs.Task.Result)
            {
                Dispatcher.Invoke(() =>
                {
                    this.ExcelIsOpenWarningDialog.IsOpen = false;
                });
                
                tcs = new TaskCompletionSource<bool>();
            }
            else
            {
                Dispatcher.Invoke(() =>
                {
                    this.ExcelIsOpenWarningDialog.IsOpen = false;
                    this._logger.OkayText = "Installation cancelled by user.";
                });
                
                tcs = new TaskCompletionSource<bool>();
                return;
            }
        } 
        
        Dispatcher.Invoke(() =>
        {
            this._logger.NewProcess("Installation of ∂Excel started.");
            this._logger.NewSubProcess("Ensuring all instances of Excel are terminated.");
        });

        CloseAllExcelInstances();

        // Remove initial (obsolete) version of dExcel.
        // Only applies to first adopters => this step can be deprecated later.
        Dispatcher.Invoke(() =>
        {
            this._logger.NewSubProcess($"Removing obsolete instances of ∂Excel.");
        });
        try
        {
            var currentAddIns =
                Directory.GetFiles(Environment.ExpandEnvironmentVariables(@"%appdata%\Microsoft\AddIns"));
            var obsoleteDExcelAddIn =
                currentAddIns.Length == 0 ?
                null : currentAddIns.First(x => x.Contains("dExcel", StringComparison.InvariantCultureIgnoreCase));
            if (obsoleteDExcelAddIn != null)
            {
                Dispatcher.Invoke(() =>
                {
                    this._logger.OkayText =
                        $"Found obsolete AddIn {Path.GetFileName(obsoleteDExcelAddIn)} in " +
                        $"[[{Path.GetDirectoryName(obsoleteDExcelAddIn)}]].";
                });
                File.Delete(obsoleteDExcelAddIn);
                Dispatcher.Invoke(() =>
                {
                    this._logger.OkayText = $"Removed obsolete AddIn [[{Path.GetFileName(obsoleteDExcelAddIn)}]].";
                });
            }
            else
            {
                Dispatcher.Invoke(() =>
                {
                    this._logger.OkayText = $"No orphaned instances of ∂Excel found.";
                });
            }
        }
        catch (Exception exception)
        {
            Dispatcher.Invoke(() =>
            {
                this._logger.ErrorText =
                    "Error removing obsolete instances of the ∂Excel add-in from " +
                    $"{Environment.ExpandEnvironmentVariables("%appdata%/Microsoft/AddIns")}.";
                this._logger.ErrorText = $"Exception message: {exception.Message}";
                this._logger.InstallationFailed();
            });
            return;
        }

        // Check if installation path exists.
        Dispatcher.Invoke(() =>
        {
            this._logger.NewSubProcess("Checking if ∂Excel installation path exists.");
        });

        if (!Directory.Exists(LocalCurrentReleasePath))
        {
            Dispatcher.Invoke(() =>
            {
                this._logger.OkayText = $"Path [[{LocalCurrentReleasePath}]] does not exist.";
            });
            
            try
            {
                Directory.CreateDirectory(LocalCurrentReleasePath);
                Dispatcher.Invoke(() =>
                {
                    this._logger.OkayText = $"Path [[{LocalCurrentReleasePath}]] created.";
                });
            }
            catch (Exception exception)
            {
                Dispatcher.Invoke(() =>
                {
                    this._logger.ErrorText = $"Path [[{LocalCurrentReleasePath}]] could not be created.";
                    this._logger.ErrorText = $"{exception.Message}";
                    this._logger.InstallationFailed();
                });
            }
        }

        Dispatcher.Invoke(() =>
        {
            this._logger.OkayText = $"Path [[{LocalCurrentReleasePath}]] already exists.";
        });

        DownloadRequiredDExcelAddInFromRemoteSource();

        // Remove previously installed version from C:\GitLab\dExcelTools\Releases\Current.
        Dispatcher.Invoke(() =>
        {
            this._logger.NewSubProcess($"Updating ∂Excel.");
            this._logger.OkayText = $"Deleting previous ∂Excel version from [[{LocalCurrentReleasePath}]].";
        });

        DirectoryInfo currentVersionDirectory = new(LocalCurrentReleasePath);
        try
        {
            foreach (FileInfo file in currentVersionDirectory.GetFiles())
            {
                file.Delete();
            }
            foreach (DirectoryInfo dir in currentVersionDirectory.GetDirectories())
            {
                dir.Delete(true);
            }
        }
        catch (Exception exception)
        {
            Dispatcher.Invoke(() =>
            {
                this._logger.ErrorText = $"Failed to delete files and folders from [[{LocalCurrentReleasePath}]].";
                this._logger.ErrorText = $"Exception message: {exception.Message}";
                this._logger.InstallationFailed();
            });
            
            return;
        }
        
        // Copy required version from 'C:\GitLab\dExcelTools\Releases\<version number>'
        // to 'C:\GitLab\dExcelTools\Releases\Current'.
        Dispatcher.Invoke(() =>
        {
            this._logger.OkayText = 
                $"Copying version {AvailableDExcelReleases.SelectedItem} of ∂Excel to [[{LocalCurrentReleasePath}]].";
        });
        try
        {
            Dispatcher.Invoke(() =>
            {
                CopyFilesRecursively(LocalReleasesPath + AvailableDExcelReleases.SelectedItem, LocalCurrentReleasePath);
            });
        }
        catch (Exception exception)
        {
            Dispatcher.Invoke(() =>
            {
                this._logger.ErrorText = 
                    $"Failed to copy version {AvailableDExcelReleases.SelectedItem} of ∂Excel to " +
                    $"[[{LocalCurrentReleasePath}]].";
                this._logger.ErrorText = $"Exception message: {exception.Message}";
                this._logger.InstallationFailed();
            });
            return;
        }

        // Create Excel application and install ∂Excel add-in.
        Dispatcher.Invoke(() =>
        {
            this._logger.OkayText = "Installing ∂Excel to Excel.";
        });
        
        // Excel is temperamental when installing an add-in via C# the very first time.
        // The first time it is best to use a VBA installer however we try the C# route first regardless.
        var failedToInstallTryVbaInstaller = true;
        
        try
        {
            var excel = new Excel.Application();
            var dExcelAdded = false;
            foreach (Excel.AddIn addIn in excel.AddIns)
            {
                if (addIn.Name.Contains("dExcel", StringComparison.InvariantCultureIgnoreCase))
                {
                    addIn.Installed = true;
                    dExcelAdded = true;
                    break;
                }
            }
            // TODO: Check if file exists
            if (!dExcelAdded)
            {
                Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
                Excel.AddIn dExcelAddIn =
                    excel.AddIns.Add(@"C:\GitLab\dExcelTools\Releases\Current\dExcel-AddIn64.xll");
                dExcelAddIn.Installed = true;
            }
            excel.Quit();
            failedToInstallTryVbaInstaller = false;
        }
        catch (Exception exception)
        {
            Dispatcher.Invoke(() =>
            {
                this._logger.WarningText = "Failed to install ∂Excel in Excel using C#-based approach.";
                this._logger.WarningText = $"Exception message: {exception.Message}";
                CloseAllExcelInstances();
            });
        }

        // Invoke the VBA installer if the C# installer failed (the case for the very first installation typically).
        try
        {
            if (failedToInstallTryVbaInstaller)
            {
                Dispatcher.Invoke(() =>
                {
                    this._logger.WarningText = "Trying to install ∂Excel using VBA.";
                    this._logger.OkayText = $"Opening Excel workbook {SharedDriveInstallationWorkbookPath}.";
                    this._logger.WarningText =
                        "If this is the first time you are installing dExcel try manually installing in Excel it from " +
                        $"[[{LocalCurrentReleasePath}]].";
                });
                    
                var excel = new Excel.Application
                {
                    Visible = true
                };
                excel.Workbooks.Open(SharedDriveInstallationWorkbookPath);
                excel.Quit();
            }
        }
        catch (Exception e)
        {
            this._logger.ErrorText = e.Message;
            this._logger.InstallationFailed();
        }

        Dispatcher.Invoke(() =>
        {
            this._logger.InstallationSucceeded();
            this.Install.IsEnabled = false;
            this.Uninstall.IsEnabled = true;
            this.CurrentDExcelVersion.Output = GetInstalledDExcelVersion();
            this.Cancel.Content = "Close";
        });
    }


    /// <summary>
    /// Copies all files, including subdirectories, from one path to path.
    /// </summary>
    /// <param name="sourcePath">Source path.</param>
    /// <param name="targetPath">Target path.</param>
    private static void CopyFilesRecursively(string sourcePath, string targetPath)
    {
        foreach (string dirPath in Directory.GetDirectories(sourcePath, "*", SearchOption.AllDirectories))
        {
            Directory.CreateDirectory(dirPath.Replace(sourcePath, targetPath));
        }

        foreach (string newPath in Directory.GetFiles(sourcePath, "*.*", SearchOption.AllDirectories))
        {
            File.Copy(newPath, newPath.Replace(sourcePath, targetPath), true);
        }
    }

    /// <summary>
    /// Deletes all files, including subdirectories, from the given directory. It does not delete the directory itself.
    /// </summary>
    /// <param name="directoryPath">The directory path.</param>
    private static void DeleteFilesRecursively(string directoryPath)
    {
        var directory = new DirectoryInfo(directoryPath);
        directory.EnumerateFiles().ToList().ForEach(f => f.Delete());
        directory.EnumerateDirectories().ToList().ForEach(d => d.Delete(true)); 
    }
    
    /// <summary>
    /// Closes the installer dialog.
    /// </summary>
    /// <param name="sender">The sender.</param>
    /// <param name="e">The RoutedEventArgs.</param>
    private void Cancel_Click(object sender, RoutedEventArgs e) => this.Close();

    /// <summary>
    /// Uninstalls ∂Excel from Excel.
    /// </summary>
    /// <param name="sender">Sender.</param>
    /// <param name="e">Routed event args.</param>
    private void Uninstall_Click(object sender, RoutedEventArgs e) => new Thread(UninstallAddIn).Start();

    /// <summary>
    /// Uninstallation process triggered by clicking the uninstall button.
    /// </summary>
    private void UninstallAddIn()
    {
        Dispatcher.Invoke(() =>
        {
            this._logger.NewProcess("Uninstalling ∂Excel from Excel.");
            this._logger.NewSubProcess("Closing all instances of Excel.");
        });
        
        CloseAllExcelInstances();
        
        try
        {
            Dispatcher.Invoke(() =>
            {
                this._logger.NewSubProcess("Opening (background) instance of Excel.");
            });
            Excel.Application excel = new();
            foreach (Excel.AddIn addIn in excel.AddIns)
            {
                if (addIn.Name.Contains("dExcel", StringComparison.InvariantCultureIgnoreCase))
                {
                    addIn.Installed = false;
                    Dispatcher.Invoke(() =>
                    {
                        this._logger.OkayText = "∂Excel uninstalled from Excel.";
                    });
                    break;
                }
            }
            excel.Quit();
        }
        catch (Exception exception)
        {
            Dispatcher.Invoke(() =>
            {
                this._logger.ErrorText = $"Exception message: {exception.Message}";
                this._logger.UninstallationFailed();
            });
            return;
        }

        try
        {
            Dispatcher.Invoke(() =>
            {
                this._logger.OkayText = $"Deleting contents of [[{LocalCurrentReleasePath}]].";
                DeleteFilesRecursively(LocalCurrentReleasePath);
            });
        }
        catch (Exception exception)
        {
            Dispatcher.Invoke(() =>
            {
                this._logger.ErrorText = $"Failed to delete ∂Excel files from [[{LocalCurrentReleasePath}]].";
                this._logger.ErrorText = exception.Message;
                this._logger.UninstallationFailed();
            });
        }

        Dispatcher.Invoke(() =>
        {
            this._logger.UninstallationSucceeded();
            this.CurrentDExcelVersion.Output = GetInstalledDExcelVersion();
            this.Uninstall.IsEnabled = false;
            this.Install.IsEnabled = true;
        });
    }

    /// <summary>
    /// Closes all instances of Excel.
    /// </summary>
    private void CloseAllExcelInstances()
    {
        try
        {
            var excelInstances = Process.GetProcessesByName("Excel");
            Dispatcher.Invoke(() =>
            {
                this._logger.OkayText = $"Excel instances found: {excelInstances.Length}";
            });
            foreach (Process excelInstance in excelInstances)
            {
                excelInstance.Kill();
                excelInstance.WaitForExit();
                excelInstance.Dispose();
            }
            Dispatcher.Invoke(() =>
            {
                this._logger.OkayText = $"All Excel instances terminated.";
            });
        }
        catch (Exception exception)
        {
            Dispatcher.Invoke(() =>
            {
                this._logger.ErrorText = $"Excel instance could not be terminated.";
                this._logger.ErrorText = $"{exception.Message}";
            });
        }
    }

    /// <summary>
    /// Event triggered by changing the selected dExcel version to install.
    /// </summary>
    /// <param name="sender">The sender.</param>
    /// <param name="e">The RoutedEventArgs.</param>
    private void AvailableDExcelReleases_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (AvailableDExcelReleases.SelectedItem != null)
        {
            Install.IsEnabled = AvailableDExcelReleases.SelectedItem.ToString() != CurrentDExcelVersion.Output;
        }
    }

    /// <summary>
    /// Launches an instance of Excel and closes the installation dialog.
    /// </summary>
    /// <param name="sender">The Sender.</param>
    /// <param name="e">The RoutedEventArgs.</param>
    private void CloseAndLaunchExcel_OnClick(object sender, RoutedEventArgs e)
    {
        try
        {
            Process.Start(@"C:\Program Files\Microsoft Office\root\Office16\excel.exe");
            this.Close();
        }
        catch (Exception exception)
        {
            this._logger.ErrorText = "Failed to launch Excel.";
            this._logger.ErrorText = exception.Message;
        }
    }

    /// <summary>
    /// Launches an instance of Excel.
    /// </summary>
    /// <param name="sender">The Sender.</param>
    /// <param name="e">The RoutedEventArgs.</param>
    private void LaunchExcel_OnClick(object sender, RoutedEventArgs e)
    {
        try
        {
            Process.Start(@"C:\Program Files\Microsoft Office\root\Office16\excel.exe");
            this._logger.OkayText = "Opening Excel.";
            this._logger.OkayText = @"[[C:\Program Files\Microsoft Office\root\Office16\excel.exe]]";
        }
        catch (Exception exception)
        {
            this._logger.ErrorText = "Failed to launch Excel.";
            this._logger.ErrorText = exception.Message;
        }
    }

    private void DownloadRequiredDExcelAddInFromRemoteSource()
    {
        Dispatcher.Invoke(() =>
        {
            if (string.Compare(
                    strA: DExcelRemoteSource.Text, 
                    strB: "Shared Drive", 
                    comparisonType: StringComparison.InvariantCultureIgnoreCase) == 0)
            {
                if (!GetAllAvailableLocalDExcelReleases().Contains(AvailableDExcelReleases.Text))
                {
                    var sourcePath = Path.Combine(SharedDriveReleasesPath, $"{AvailableDExcelReleases.Text}.zip");
                    var targetPath = Path.Combine(LocalReleasesPath, $"{AvailableDExcelReleases.Text}");
                    var zipOutputPath = Path.Combine(LocalReleasesPath, $"{AvailableDExcelReleases.Text}");
                    File.Copy(sourcePath, targetPath, true);
                    ZipFile.ExtractToDirectory(targetPath, Path.GetDirectoryName(targetPath) ?? string.Empty);
                    File.Delete(targetPath);
                }
            }
        });
    }
}
