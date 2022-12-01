namespace dExcel.WPF;

using LiveChartsCore;
using LiveChartsCore.SkiaSharpView;

public class ViewModel
{
    public ISeries[] Series { get; set; } 
        = {
            new LineSeries<double>
            {
                Values = new double[] { 2, 1, 3, 5, 3, 4, 6 },
                Fill = null
            }
        }; 
}
