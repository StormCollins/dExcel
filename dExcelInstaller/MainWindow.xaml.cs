namespace dExcelInstaller;

using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Security.Principal;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using Excel = Microsoft.Office.Interop.Excel;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    private const string VersionsPath = @"C:\GitLab\dExcelTools\Versions\";
    private const string CurrentVersionPath = @"C:\GitLab\dExcelTools\Versions\Current\";
    private const string Dll = "dExcel.dll";
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
        InstallerVersion.Text = $"{installerVersion?.Major}.{installerVersion?.Minor}";
        
        CurrentDExcelVersion.Text = GetCurrentDExcelVersion();
        if (string.Compare(
                strA: GetCurrentDExcelVersion(), 
                strB: "Not Installed", 
                comparisonType: StringComparison.InvariantCultureIgnoreCase) == 0)
        {
            this.Uninstall.IsEnabled = false;
        }
        
        
        AdminRights.Text = IsAdministrator().ToString();
        if (!IsAdministrator())
        {
            this._logger.WarningText =
                "The user is not an admin on this machine. It may limit installation functionality.";
        }

        if (NetworkUtils.GetConnectionStatus())
        {
            _vpnConnectionStatus = true;
            this.ConnectionStatus.Source =
                new BitmapImage(
                    new Uri(@"pack://application:,,,/resources/icons/connection-status-green.ico",
                    UriKind.Absolute));
            this.DockPanelConnectionStatus.ToolTip = "You are connected to the VPN.";
        }
        else
        {
            _vpnConnectionStatus = false;
            this.ConnectionStatus.Source =
                new BitmapImage(
                    new Uri(
                        uriString: @"pack://application:,,,/resources/icons/connection-status-amber.ico",
                        uriKind: UriKind.Absolute));
            this._logger.WarningText =
                "User not connected to VPN. VPN is required to check for latest versions of ∂Excel on GitLab.";
            this.DockPanelConnectionStatus.ToolTip = "You are not connected to the VPN.";
        }

        NetworkChange.NetworkAddressChanged += ConnectionStatusChangedCallback;

        // TODO: Get new versions from GitLab before this step.
        AvailableDExcelVersions.ItemsSource = GetLocalAvailableDExcelVersions();
        AvailableDExcelVersions.SelectedIndex = 0;
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
                    this._logger.OkayText =
                        "Connection to VPN established. Checking for latest versions of ∂Excel on GitLab.";
                    this.DockPanelConnectionStatus.ToolTip = "You are connected to the VPN.";
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
                    this._logger.WarningText =
                        "User not connected to VPN. VPN is required to check for latest versions of ∂Excel on GitLab.";
                    this.DockPanelConnectionStatus.ToolTip = "You are not connected to the VPN.";
                });
            }
        }
    }

    /// <summary>
    /// Gets all versions of dExcel already copied to the user's local machine.
    /// </summary>
    /// <returns>Available local versions of dExcel</returns>
    private static IEnumerable<string?> GetLocalAvailableDExcelVersions()
    {
        return Directory
            .GetDirectories(VersionsPath)
            .Where(x => Regex.IsMatch(x, @"\d+(.\d+)"))
            .Select(Path.GetFileName)
            .Reverse();
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
    private static string GetCurrentDExcelVersion()
    {
        if (File.Exists(CurrentVersionPath + @"\" + Dll))
        {
            var currentDExcelVersion = AssemblyName.GetAssemblyName(CurrentVersionPath + @"\" + Dll).Version;
            return $"{currentDExcelVersion?.Major}.{currentDExcelVersion?.Minor}";
        }

        return "Not Installed";
    }


    /// <summary>
    /// Installs the specified version of the dExcel AddIn to Excel.
    /// </summary>
    private void InstallAddIn()
    {
        Dispatcher.Invoke(() =>
        {
            this._logger.NewProcess("Installation of ∂Excel started.");
            this._logger.NewSubProcess($"Ensuring Excel is closed.");
        });

        // Ensure Excel is closed.
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
                        $"{Path.GetDirectoryName(obsoleteDExcelAddIn)}.";
                });
                File.Delete(obsoleteDExcelAddIn);
                Dispatcher.Invoke(() =>
                {
                    this._logger.OkayText = $"Removed obsolete AddIn {Path.GetFileName(obsoleteDExcelAddIn)}.";
                });
            }
            else
            {
                Dispatcher.Invoke(() =>
                {
                    this._logger.OkayText = $"No obsolete instances of ∂Excel found.";
                });
            }
        }
        catch (Exception exception)
        {
            Dispatcher.Invoke(() =>
            {
                this._logger.ErrorText =
                    $"Error removing obsolete instances of ∂Excel from " +
                    $"{Environment.ExpandEnvironmentVariables("%appdata%/Microsoft/AddIns")}.";
                this._logger.ErrorText = exception.Message;
                this._logger.InstallationFailed();
            });
            return;
        }

        // Check if installation path exists.
        Dispatcher.Invoke(() =>
        {
            this._logger.NewSubProcess($"Checking if {VersionsPath} exists.");
        });

        if (!Directory.Exists(VersionsPath))
        {
            Dispatcher.Invoke(() =>
            {
                this._logger.OkayText = $"Path {VersionsPath} does not exist.";
            });
            
            try
            {
                Directory.CreateDirectory(VersionsPath);
                Dispatcher.Invoke(() =>
                {
                    this._logger.OkayText = $"Path {VersionsPath} created.";
                });
            }
            catch (Exception exception)
            {
                Dispatcher.Invoke(() =>
                {
                    this._logger.ErrorText = $"Path {VersionsPath} could not be created.";
                    this._logger.ErrorText = $"{exception.Message}";
                    this._logger.InstallationFailed();
                });
            }
            return;
        }

        Dispatcher.Invoke(() =>
        {
            this._logger.OkayText = $"Path {VersionsPath} already exists.";
        });

        // Download addin from GitLab and copy to installation path.


        // Remove previously installed version from C:\GitLab\dExcelTools\Versions\Current.
        Dispatcher.Invoke(() =>
        {
            this._logger.NewSubProcess($"Updating ∂Excel.");
            this._logger.OkayText = $"Deleting previous ∂Excel version from {CurrentVersionPath}.";
        });
        DirectoryInfo currentVersionDirectory = new(CurrentVersionPath);
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
                this._logger.ErrorText = $"Failed to delete files and folders from {CurrentVersionPath}.";
                this._logger.ErrorText = exception.Message;
                this._logger.InstallationFailed();
            });
            return;
        }

        // Copy required version from 'C:\GitLab\dExcelTools\Versions\<version number>'
        // to 'C:\GitLab\dExcelTools\Versions\Current'.
        Dispatcher.Invoke(() =>
        {
            this._logger.OkayText = $"Copying version {AvailableDExcelVersions.SelectedItem} of ∂Excel to {CurrentVersionPath}.";
        });
        try
        {
            Dispatcher.Invoke(() =>
            {
                CopyFilesRecursively(VersionsPath + AvailableDExcelVersions.SelectedItem, CurrentVersionPath);
            });
        }
        catch (Exception exception)
        {
            Dispatcher.Invoke(() =>
            {
                this._logger.ErrorText = $"Failed to copy version {AvailableDExcelVersions.SelectedItem} of ∂Excel to {CurrentVersionPath}.";
                this._logger.ErrorText = $"{exception.Message}";
                this._logger.InstallationFailed();
            });
            return;
        }

        // Create Excel application and ∂Excel install addin.
        Dispatcher.Invoke(() =>
        {
            this._logger.OkayText = $"Installing ∂Excel in Excel.";
        });

        try
        {
            Excel.Application excel = new();
            var dExcelInstalled = false;
            foreach (Excel.AddIn addIn in excel.AddIns)
            {
                if (addIn.Name.Contains("dExcel"))
                {
                    addIn.Installed = true;
                    dExcelInstalled = true;
                    break;
                }
            }

            if (!dExcelInstalled)
            {
                Excel.AddIn dExcelAddIn =
                    excel.AddIns.Add(@"C:\GitLab\dExcelTools\Versions\Current\dExcel-AddIn64.xll");
                dExcelAddIn.Installed = true;
            }
            excel.Quit();
        }
        catch (Exception exception)
        {
            Dispatcher.Invoke(() =>
            {
                this._logger.ErrorText = $"Failed to install ∂Excel in Excel.";
                this._logger.ErrorText = $"{exception.Message}";
                this._logger.InstallationFailed();
            });
            CloseAllExcelInstances();
            Dispatcher.Invoke(() =>
            {
                this._logger.InstallationFailed();
            });
            return;
        }

        Dispatcher.Invoke(() =>
        {
            this._logger.InstallationSucceeded();
            this.Install.IsEnabled = false;
            this.Uninstall.IsEnabled = true;
            this.CurrentDExcelVersion.Text = GetCurrentDExcelVersion();
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
                this._logger.NewSubProcess("Opening instance of Excel.");
            });
            Excel.Application excel = new();
            foreach (Excel.AddIn addIn in excel.AddIns)
            {
                if (addIn.Name.Contains("dExcel"))
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
                this._logger.ErrorText = exception.Message;
                this._logger.UninstallationFailed();
            });
            return;
        }

        try
        {
            Dispatcher.Invoke(() =>
            {
                this._logger.OkayText = $"Deleting contents of {CurrentVersionPath}.";
                DeleteFilesRecursively(CurrentVersionPath);
            });
        }
        catch (Exception exception)
        {
            Dispatcher.Invoke(() =>
            {
                this._logger.ErrorText = $"Failed to delete ∂Excel files from {CurrentVersionPath}.";
                this._logger.ErrorText = exception.Message;
                this._logger.UninstallationFailed();
            });
        }

        Dispatcher.Invoke(() =>
        {
            this._logger.UninstallationSucceeded();
            this.CurrentDExcelVersion.Text = GetCurrentDExcelVersion();
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
                this._logger.OkayText = $"{excelInstances.Length} Excel instances found.";
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
                this._logger.ErrorText = $"Excel process could not be terminated.";
                this._logger.ErrorText = $"{exception.Message}";
            });
        }
    }

    /// <summary>
    /// Event triggered by changing the selected dExcel version to install.
    /// </summary>
    /// <param name="sender">The sender.</param>
    /// <param name="e">The RoutedEventArgs.</param>
    private void AvailableDExcelVersions_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        Install.IsEnabled = AvailableDExcelVersions.SelectedItem.ToString() != CurrentDExcelVersion.Text;
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
}
