namespace dExcel;

using ExcelDna.Integration;
using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;

/// <summary>
/// The formatting settings to apply to a table.
/// </summary>
/// <param name="IsVertical">True if the table is vertical (i.e. has column headers) else false if horizontal 
/// (i.e. has row headers).</param>
/// <param name="HasTwoHeaders">True if the table has two header rows.</param>
public readonly record struct FormattingSettings(int columnHeaderCount, int rowHeaderCount);

/// <summary>
/// Interaction logic for TableFormatter.xaml which allows users to quickly select and apply the format for a selected
/// table.
/// </summary>
public partial class TableFormatter : Window
{
    private static TableFormatter? _instance;

    public FormattingSettings? FormatSettings { get; set; }

    /// <summary>
    /// Creates an instance of <see cref="TableFormatter"/> using the Singleton pattern.
    /// </summary>
    public static TableFormatter Instance
    {
        get
        {
            if (_instance == null)
            {
                _instance = new TableFormatter();
            }
            return _instance;
        }
    }

    /// <summary>
    /// Creates an instance of <see cref="TableFormatter"/>.
    /// </summary>
    private TableFormatter()
    {
        InitializeComponent();
        Closing += TableFormatter_Closing;
    }

    void OnLoad(object sender, RoutedEventArgs e)
    {
        HasZeroColumnHeaders.IsChecked = false;
        HasOneColumnHeader.IsChecked = true;
        HasTwoColumnHeaders.IsChecked = false;
        HasZeroRowHeaders.IsChecked = true;
        HasOneRowHeader.IsChecked = false;
        HasTwoRowHeaders.IsChecked = false;
    }

        /// <summary>
        /// Event called when TableFormatter WPF form closes.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">Event args.</param>
        private void TableFormatter_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        _instance = null;
    }

    /// <summary>
    /// Event called when TableFormatter 'Close' button or red 'X' is clicked.
    /// </summary>
    /// <param name="sender">The sender.</param>
    /// <param name="e">Event args.</param>
    private void CloseTableFormatter(object sender, RoutedEventArgs e)
    {
        this.FormatSettings = null;
        this.Close();
    }

    private void FormatTable_Click(object sender, RoutedEventArgs e)
    {
        int columnHeaderCount = (bool)HasZeroColumnHeaders.IsChecked ? 0 : (bool)HasOneColumnHeader.IsChecked ? 1 : 2;
        int rowHeaderCount = (bool)HasZeroRowHeaders.IsChecked ? 0 : (bool)HasOneRowHeader.IsChecked ? 1 : 2;
        this.FormatSettings = new FormattingSettings(columnHeaderCount, rowHeaderCount);
        this.Close();
    }


    private void Headers_Checked(object sender, RoutedEventArgs e)
    {
        if ((bool)HasZeroColumnHeaders.IsChecked && (bool)HasOneRowHeader.IsChecked)
        {
            var xllPath = Path.GetDirectoryName(ExcelDnaUtil.XllPath);
            Example.Source = new BitmapImage(new Uri(xllPath + @"\resources\icons\table-formatting-row-1.png"));
        }
        else if ((bool)HasZeroColumnHeaders.IsChecked && (bool)HasTwoRowHeaders.IsChecked)
        {
            var xllPath = Path.GetDirectoryName(ExcelDnaUtil.XllPath);
            Example.Source = new BitmapImage(new Uri(xllPath + @"\resources\icons\table-formatting-row-2.png"));
        }
        else if ((bool)HasOneColumnHeader.IsChecked && (bool)HasZeroRowHeaders.IsChecked)
        {
            var xllPath = Path.GetDirectoryName(ExcelDnaUtil.XllPath);
            Example.Source = new BitmapImage(new Uri(xllPath + @"\resources\icons\table-formatting-column-1.png"));
        }
        else if ((bool)HasOneColumnHeader.IsChecked && (bool)HasOneRowHeader.IsChecked)
        {
            var xllPath = Path.GetDirectoryName(ExcelDnaUtil.XllPath);
            Example.Source = new BitmapImage(new Uri(xllPath + @"\resources\icons\table-formatting-column-1-row-1.png"));
        }
        else if ((bool)HasOneColumnHeader.IsChecked && (bool)HasTwoRowHeaders.IsChecked)
        {
            var xllPath = Path.GetDirectoryName(ExcelDnaUtil.XllPath);
            Example.Source = new BitmapImage(new Uri(xllPath + @"\resources\icons\table-formatting-column-1-row2.png"));
        }
        else if ((bool)HasTwoColumnHeaders.IsChecked && (bool)HasZeroRowHeaders.IsChecked)
        {
            var xllPath = Path.GetDirectoryName(ExcelDnaUtil.XllPath);
            Example.Source = new BitmapImage(new Uri(xllPath + @"\resources\icons\table-formatting-column-2.png"));
        }
        else if ((bool)HasTwoColumnHeaders.IsChecked && (bool)HasOneRowHeader.IsChecked)
        {
            var xllPath = Path.GetDirectoryName(ExcelDnaUtil.XllPath);
            Example.Source = new BitmapImage(new Uri(xllPath + @"\resources\icons\table-formatting-column-2-row-1.png"));
        }
        else if ((bool)HasTwoColumnHeaders.IsChecked && (bool)HasTwoRowHeaders.IsChecked)
        {
            var xllPath = Path.GetDirectoryName(ExcelDnaUtil.XllPath);
            Example.Source = new BitmapImage(new Uri(xllPath + @"\resources\icons\table-formatting-column-2-row-2.png"));
        }

        if ((bool)HasZeroColumnHeaders.IsChecked)
        {
            HasZeroRowHeaders.IsEnabled = false;
        }
        else
        {
            HasZeroRowHeaders.IsEnabled = true;
        }

        if ((bool)HasZeroRowHeaders.IsChecked)
        {
            HasZeroColumnHeaders.IsEnabled = false;
        }
        else
        {
            HasZeroColumnHeaders.IsEnabled = true;
        }
    }
}
