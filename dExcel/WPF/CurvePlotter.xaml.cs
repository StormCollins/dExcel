using LiveChartsCore.Kernel.Sketches;

namespace dExcel.WPF;

using dExcel.InterestRates;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;
using LiveChartsCore.Defaults;
using LiveChartsCore.SkiaSharpView.Painting;
using LiveChartsCore.SkiaSharpView;
using LiveChartsCore;
using SkiaSharp;
using System.Collections.ObjectModel;
using System.Windows;

/// <summary>
/// Interaction logic for TableFormatter.xaml which allows users to quickly select and apply the format for a selected
/// table.
/// </summary>
public partial class CurvePlotter: Window
{
    private static CurvePlotter? _instance;

    /// <summary>
    /// Creates an instance of <see cref="TableFormatter"/> using the Singleton pattern.
    /// </summary>
    public static CurvePlotter Instance => _instance ??= new CurvePlotter();

    /// <summary>
    /// Creates an instance of <see cref="CurvePlotter"/>.
    /// </summary>
    private CurvePlotter()
    {
        InitializeComponent();
        Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
        CurveDetails curveDetails = CurveUtils.GetCurveDetails(xlApp.Selection.Value2);
        ObservableCollection<DateTimePoint> values = new();
        for (int i = 0; i < curveDetails.DiscountFactorDates?.Count; i++)
        {
            values.Add(new DateTimePoint(curveDetails.DiscountFactorDates[i].ToDateTime(), curveDetails.DiscountFactors?[i]));
        }
        
        ChartArea.Series = new ObservableCollection<ISeries>
        {
            new LineSeries<DateTimePoint>
            {
                Fill = null,
                GeometryFill = null,
                GeometrySize = 7,
                GeometryStroke = new SolidColorPaint(SKColor.Parse("#86BC25")),
                Stroke = new SolidColorPaint(SKColor.Parse("#86BC25")) { StrokeThickness = 3 },
                Values = values,
            },
        };

        ChartArea.XAxes = new Axis[]
        {
            new Axis
            {
                Labeler = x => new DateTime((long)x).ToString("yyyy-MM-dd"),
                LabelsPaint = new SolidColorPaint(SKColors.White),
                LabelsRotation = 15,
                Name = "Dates",
                NamePaint = new SolidColorPaint(SKColors.White),
            }
        };
        
        Closing += CurvePlotter_Closing;
    }

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
        this.Close();
    }

    private void btnClose_Click(object sender, RoutedEventArgs e)
    {
        this.CloseCurvePlotter(sender, e);
    }
}
