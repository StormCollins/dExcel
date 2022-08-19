namespace dExcelWpf;

using System.ComponentModel;
using System.Windows.Controls;

/// <summary>
/// Specifies a label with a helpful tooltip.
/// </summary>
public partial class InformationalLabel : UserControl 
{
    public InformationalLabel()
    {
        InitializeComponent();
        this.DataContext = this;
    }
    
    public string? Label { get; set; }
    
    public string? Tip { get; set; }
}
