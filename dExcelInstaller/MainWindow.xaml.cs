namespace dExcelInstaller;

using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media.Imaging;
using Excel = Microsoft.Office.Interop.Excel;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    private const string VersionsPath = @"C:\GitLab\dExcelTools\Versions\";
    private const string CurrentVersionPath = @"C:\GitLab\dExcelTools\Versions\Current\";
    private const string dExcelVersion = "1.1";
    private static bool _connectionStatus;

    // What does it mean to register an addin?
    // Difference between COM addin?
    // What is C API for Excel?
    // stackoverflow.com/questions/363377/how-do-i-run-a-simple-bit-of-code-in-a-new-thread
    public Logger logger;

    public MainWindow()
    {
        InitializeComponent();
        dExcelIcon.Source = new BitmapImage(new Uri(@"pack://application:,,,/resources/icons/dExcel.png", UriKind.Absolute));
        var installerVersion = Assembly.GetEntryAssembly()?.GetName().Version;
        InstallerVersion.Text = $"{installerVersion?.Major}.{installerVersion?.Minor}";
        var currentDExcelVersion = AssemblyName.GetAssemblyName(CurrentVersionPath + @"\dExcel.dll").Version;
        CurrentDExcelVersion.Text = $"{currentDExcelVersion?.Major}.{currentDExcelVersion?.Minor}";
        
        if (NetworkUtils.GetConnectionStatus())
        {
            _connectionStatus = true;
            ConnectionStatus.Source = new BitmapImage(new Uri(@"pack://application:,,,/resources/icons/connection-status-green.ico", UriKind.Absolute));
        }
        else
        {
            _connectionStatus = false;
            ConnectionStatus.Source = new BitmapImage(new Uri(@"pack://application:,,,/resources/icons/connection-status-amber.ico", UriKind.Absolute));
        }
        NetworkChange.NetworkAddressChanged += ConnectionStatusChangedCallback;

        // TODO: Get new versions from GitLab before this step.
        AvailableDExcelVersions.ItemsSource = GetAvailableVersions();
        AvailableDExcelVersions.SelectedIndex = 0;
        this.logger = new Logger(LogWindow);
    }


    public void ConnectionStatusChangedCallback(object sender, EventArgs e)
    {
        if (NetworkUtils.GetConnectionStatus() != _connectionStatus)
        {
            _connectionStatus = !_connectionStatus;
            if (NetworkUtils.GetConnectionStatus())
            {
                Dispatcher.Invoke(() =>
                    ConnectionStatus.Source = new BitmapImage(new Uri(@"pack://application:,,,/resources/icons/connection-status-green.ico", UriKind.Absolute)));
            }
            else
            {
                Dispatcher.Invoke(() =>
                    ConnectionStatus.Source = new BitmapImage(new Uri(@"pack://application:,,,/resources/icons/connection-status-amber.ico", UriKind.Absolute)));
            }
        }
    }

    /// <summary>
    /// Gets all versions of dExcel already copied to the user's local machine.
    /// </summary>
    /// <returns>Available local versions of dExcel</returns>
    private static IEnumerable<string> GetAvailableVersions()
    {
        return Directory
            .GetDirectories(VersionsPath)
            .Where(x => Regex.IsMatch(x, @"\d+(.\d+)"))
            .Select(x => Path.GetFileName(x))
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


    public void InstallAddIn()
    {
        Dispatcher.Invoke(() =>
        {
            this.logger.NewProcess("Installation of ∂Excel started.");
            this.logger.NewSubProcess($"Ensuring Excel is closed.");
        });

        // Ensure Excel is closed.
        try
        {
            Process[] excelInstances = Process.GetProcessesByName("Excel");
            Dispatcher.Invoke(() =>
            {
                this.logger.OkayText = $"{excelInstances.Length} Excel instances found.";
            });
            foreach (Process excelInstance in excelInstances)
            {
                excelInstance.Kill();
                excelInstance.WaitForExit();
                excelInstance.Dispose();
            }
            Dispatcher.Invoke(() =>
            {
                this.logger.OkayText = $"All Excel instances terminated.";
            });
        }
        catch (Exception exception)
        {
            Dispatcher.Invoke(() =>
            {
                this.logger.ErrorText = $"Excel process could not be terminated.";
                this.logger.ErrorText = $"{exception.Message}";
            });
        }

        // Remove the initial (obsolete) version of dExcel.
        // This would only apply to first adopters so this step can be deprecated later.
        Dispatcher.Invoke(() =>
        {
            this.logger.NewSubProcess($"Removing obsolete instances of ∂Excel.");
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
                    this.logger.OkayText = $"Found obsolete AddIn {Path.GetFileName(obsoleteDExcelAddIn)} "
                    + $"in {Path.GetDirectoryName(obsoleteDExcelAddIn)}.";
                });
                File.Delete(obsoleteDExcelAddIn);
                Dispatcher.Invoke(() =>
                {
                    this.logger.OkayText = $"Removed obsolete AddIn {Path.GetFileName(obsoleteDExcelAddIn)}.";
                });
            }
        }
        catch (Exception exception)
        {
            Dispatcher.Invoke(() =>
            {
                this.logger.ErrorText = $"Error removing obsolete instances of ∂Excel from " +
                $"{Environment.ExpandEnvironmentVariables("%appdata%/Microsoft/AddIns")}";
                this.logger.ErrorText = exception.Message;
            });
        }
        Dispatcher.Invoke(() =>
        {
            this.logger.DashedHorizontalLine();
        });

        // Check if installation path exists.
        Dispatcher.Invoke(() =>
        {
        });

        if (!Directory.Exists(VersionsPath))
        {
            Dispatcher.Invoke(() =>
            {
                this.logger.OkayText = $"Path {VersionsPath} does not exist.";
            });
            
            try
            {
                Directory.CreateDirectory(VersionsPath);
                Dispatcher.Invoke(() =>
                {
                    this.logger.OkayText = $"Path {VersionsPath} created.";
                });
            }
            catch (Exception exception)
            {
                this.logger.ErrorText = $"Path {VersionsPath} could not be created.";
                this.logger.ErrorText = $"{exception.Message}";
            }
        }
        else
        {
            Dispatcher.Invoke(() =>
            {
                this.logger.OkayText = $"Path {VersionsPath} already exists.";
            });
        }
        Dispatcher.Invoke(() =>
        {
            this.logger.DashedHorizontalLine();
        });

        // Download addin from GitLab and copy to installation path.


        // Remove previous version from C:\GitLab\dExcelTools\Versions\Current.
        Dispatcher.Invoke(() =>
        {
            this.logger.NewSubProcess($"Updating ∂Excel.");
            this.logger.OkayText = $"Deleting previous ∂Excel version from {CurrentVersionPath}.";
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
                this.logger.ErrorText = $"Failed to delete files and folders from {CurrentVersionPath}.";
                this.logger.ErrorText = exception.Message;
            }); 
        }

        // Copy required version from 'C:\GitLab\dExcelTools\Versions\<version number>' to
        // 'C:\GitLab\dExcelTools\Versions\Current'.
        Dispatcher.Invoke(() =>
        {
            logger.OkayText = $"Copying new version of ∂Excel to {CurrentVersionPath}.";
        });
        try
        {
            CopyFilesRecursively(VersionsPath + dExcelVersion, CurrentVersionPath);
        }
        catch (Exception exception)
        {
            Dispatcher.Invoke(() =>
            {
                this.logger.ErrorText = $"Failed to copy new version of ∂Excel to {CurrentVersionPath}.";
                this.logger.ErrorText = $"{exception.Message}";
            });
        }


        // Create Excel application and ∂Excel install addin.
        Excel.Application excel = new();
        bool dExcelInstalled = false;
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

        Dispatcher.Invoke(() =>
        {
            Install.IsEnabled = false;
            Cancel.Content = "Close";
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

    private void Cancel_Click(object sender, RoutedEventArgs e) => this.Close();

    /// <summary>
    /// Uninstalls ∂Excel from Excel.
    /// </summary>
    /// <param name="sender">Sender.</param>
    /// <param name="e">Routed event args.</param>
    private void Uninstall_Click(object sender, RoutedEventArgs e)
    {
        Excel.Application excel = new();
        foreach (Excel.AddIn addIn in excel.AddIns)
        {
            if (addIn.Name.Contains("dExcel"))
            {
                addIn.Installed = false;
                break;
            }
        }
        excel.Quit();
    }
}
