namespace dExcel.WPF;

using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using ExcelDna.Integration;
using FuzzySharp;
using FuzzySharp.Extractor;
using LiveChartsCore;
using LiveChartsCore.SkiaSharpView;
using LiveChartsCore.SkiaSharpView.Painting;

/// <summary>
/// Interaction logic for FunctionSearch.xaml
/// </summary>
public partial class CurvePlotter : Window
{
    public CurvePlotter()
    {
        InitializeComponent();
        PlotArea.Series = new ISeries[]
        {
            new LineSeries<double>
            {
                Values = new double[] {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11},
            }
        };
    }
    
    public ISeries[] Series { get; set; } 
        = new ISeries[]
        {
            new LineSeries<double>
            {
                Values = new double[] { 2, 1, 3, 5, 3, 4, 6 },
                Fill = null
            }
        };
}
