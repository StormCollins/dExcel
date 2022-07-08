namespace dExcel;

using System;
using System.Diagnostics;
using System.Windows.Navigation;
using System.Net.NetworkInformation;
using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;
using ExcelDna.Integration;

public readonly record struct FormattingSettings(bool IsVertical, bool HasSecondaryHeader);

/// <summary>
/// Interaction logic for TableFormatter.xaml which allows users to quickly select and apply the format for a selected
/// table.
/// </summary>
public partial class TableFormatter : Window
{
    private static TableFormatter? _instance;

    public FormattingSettings? FormatSettings { get; set; }

    /// <summary>
    /// Creates an instance of TableFormatter using the Singleton pattern.
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
    /// Creates an instance of TableFormatter.
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
    /// Event called when TableFormatter close button is clicked.
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
}
