namespace dExcel;

using System.Windows;

/// <summary>
/// The formatting settings to apply to a table.
/// </summary>
/// <param name="IsVertical">True if the table is vertical (i.e. has column headers) else false if horizontal 
/// (i.e. has row headers).</param>
/// <param name="HasTwoHeaders">True if the table has two header rows.</param>
public readonly record struct FormattingSettings(bool IsVertical, bool HasTwoHeaders);

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
        this.FormatSettings
            = new FormattingSettings((bool)VerticallyAlignedTable.IsChecked, (bool)chkBxSecondaryHeader.IsChecked);
        this.Close();
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void HorizontallyAlignedTable_Checked(object sender, RoutedEventArgs e)
    {
        if (chkBxSecondaryHeader != null)
        {
            chkBxSecondaryHeader.Content = "Two Header Columns";
        }
    }

    private void VerticallyAlignedTable_Checked(object sender, RoutedEventArgs e)
    {
        if (chkBxSecondaryHeader != null)
        {
            chkBxSecondaryHeader.Content = "Two Header Rows";
        }
    }
}
