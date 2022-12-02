using LiveChartsCore.SkiaSharpView;

namespace dExcel.WPF;

using System.Windows;
using LiveChartsCore;
using SkiaSharp.Views.WPF;

/// <summary>
/// Interaction logic for TableFormatter.xaml which allows users to quickly select and apply the format for a selected
/// table.
/// </summary>
public partial class CurvePlotter: Window
{
    private static CurvePlotter? _instance;

    public FormattingSettings? FormatSettings { get; set; }

    /// <summary>
    /// Creates an instance of <see cref="TableFormatter"/> using the Singleton pattern.
    /// </summary>
    public static CurvePlotter Instance
    {
        get
        {
            if (_instance == null)
            {
                _instance = new CurvePlotter();
            }
            return _instance;
        }
    }

    /// <summary>
    /// Creates an instance of <see cref="CurvePlotter"/>.
    /// </summary>
    private CurvePlotter()
    {
        InitializeComponent();

        // SkiaSharp.Views.WPF.SKElement skElement = new SKElement();
        ChartArea.Series = new ISeries[]
        {
            new LineSeries<double>
            {
                Values = new double[] { 2, 1, 3, 5, 3, 4, 6 }
            }
        };
        Closing += CurvePlotter_Closing;
    }

    // void OnLoad(object sender, RoutedEventArgs e)
    // {
        // HasZeroColumnHeaders.IsChecked = false;
        // HasOneColumnHeader.IsChecked = true;
        // HasTwoColumnHeaders.IsChecked = false;
        // HasZeroRowHeaders.IsChecked = true;
        // HasOneRowHeader.IsChecked = false;
        // HasTwoRowHeaders.IsChecked = false;
    // }

    /// <summary>
    /// Event called when TableFormatter WPF form closes.
    /// </summary>
    /// <param name="sender">The sender.</param>
    /// <param name="e">Event args.</param>
    private void CurvePlotter_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        _instance = null;
    }

    /// <summary>
    /// Event called when TableFormatter 'Close' button or red 'X' is clicked.
    /// </summary>
    /// <param name="sender">The sender.</param>
    /// <param name="e">Event args.</param>
    private void CloseCurvePlotter(object sender, RoutedEventArgs e)
    {
        // this.FormatSettings = null;
        this.Close();
    }
}
