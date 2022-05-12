namespace dExcelInstaller;

using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Media.Imaging;
using Excel = Microsoft.Office.Interop.Excel;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    public Logger logger;

    public MainWindow()
    {
        InitializeComponent();
        this.Icon = dExcelIcon.Source = new BitmapImage(new Uri(Directory.GetCurrentDirectory() + @"\resources\icons\dXL-logo-extra-small.ico"));
        dExcelIcon.Source = new BitmapImage(new Uri(Directory.GetCurrentDirectory() + @"\resources\icons\dXL-logo.ico"));
        logger = new Logger(LogWindow);
    }

    /// <summary>
    /// Installation process triggered by clicking the install button.
    /// </summary>
    /// <param name="sender">The sender.</param>
    /// <param name="e">The routed events.</param>
    private void Install_Click(object sender, RoutedEventArgs e)
    {
        this.logger.NewProcess("Installation of ∂Excel started.");

        // Ensure Excel is closed.
        this.logger.OkayText = $"Checking that all instances of Excel are closed.";
        try
        {
            Process[] workers = Process.GetProcessesByName("Excel");
            this.logger.OkayText = $"{workers.Length} Excel instances found.";
            foreach (Process worker in workers)
            {
                worker.Kill();
                worker.WaitForExit();
                worker.Dispose();
            }
            this.logger.OkayText = $"All Excel instances terminated.";
        }
        catch (Exception exception)
        {
            this.logger.ErrorText = $"Excel process could not be terminated.";
            this.logger.ErrorText = $"{exception.Message}";
        }
        this.logger.DashedHorizontalLine();

        // Remove the initial version of dExcel - this would only apply to first adopters 
        // and this step can be deprecated later.
        this.logger.OkayText = $"Checking for obsolete instances of ∂Excel.";
        try
        {
            var currentAddIns = Directory.GetFiles(Environment.ExpandEnvironmentVariables(@"%appdata%\Microsoft\AddIns"));
            var obsoleteDExcelAddIn = currentAddIns.Length == 0 ? null : currentAddIns.First(x => x.Contains("dExcel", StringComparison.InvariantCultureIgnoreCase));
            if (obsoleteDExcelAddIn != null)
            {
                this.logger.OkayText = $"Found obsolete AddIn {Path.GetFileName(obsoleteDExcelAddIn)} "
                    + $"in {Path.GetDirectoryName(obsoleteDExcelAddIn)}.";
                File.Delete(obsoleteDExcelAddIn);
                this.logger.OkayText = $"Deleted obsolete AddIn {Path.GetFileName(obsoleteDExcelAddIn)}.";
            }
        }
        catch (Exception exception)
        {
            this.logger.ErrorText = $"Error deleting obsolete instances of ∂Excel from " +
                $"{Environment.ExpandEnvironmentVariables("%appdata%/Microsoft/AddIns")}"; 
            this.logger.ErrorText = exception.Message; 
        }
        this.logger.DashedHorizontalLine();

        // Check if installation path exists.
        var versionsPath = @"C:\GitLab\dExcelTools\Versions";
        if (!Directory.Exists(versionsPath))
        {
            this.logger.OkayText = $"Path {versionsPath} does not exist.";
            try
            {
                Directory.CreateDirectory(versionsPath);
                this.logger.OkayText = $"Path {versionsPath} created.";
            }
            catch (Exception exception)
            {
                this.logger.ErrorText = $"Path {versionsPath} could not be created.";
                this.logger.ErrorText = $"{exception.Message}";
            }
        }
        else
        {
            logger.OkayText = $"Path {versionsPath} already exists.";
        }
        this.logger.DashedHorizontalLine();

        // Download addin from GitLab and copy to installation path.

        // Open Excel and install the addin.
        // stackoverflow.com/questions/363377/how-do-i-run-a-simple-bit-of-code-in-a-new-thread
        logger.OkayText = $"Opening Excel to install ∂Excel add in.";
        try
        { 
            Thread t = new Thread(InstalldExcelAddIn);
            t.Start();
            
        }
        catch (Exception)
        {
            throw;
        }
    }

    static void InstalldExcelAddIn()
    {
        Excel.Application excel = new Excel.Application();
        Excel.Workbook wb = excel.Workbooks.Open(@"C:\GitLab\dExcelTools\dExcel\dExcel\resources\workbooks\Testing.xlsm");
        Excel.AddIn dExcelAddIn = excel.AddIns.Add(@"C:\GitLab\dExcelTools\Versions\0.0\dExcel-AddIn64.xll");
        dExcelAddIn.Installed = true;
        //excel.RegisterXLL(@"C:\GitLab\dExcelTools\Versions\0.0\dExcel - AddIn64.xll");
        excel.Visible = true;
    }

    private void Cancel_Click(object sender, RoutedEventArgs e) => this.Close();
}
