namespace dExcelWpf;

using System.ComponentModel;

/// <summary>
/// Specifies a label with a helpful tooltip.
/// </summary>
public partial class InformationalLabelWithOutput : INotifyPropertyChanged
{
    public InformationalLabelWithOutput()
    {
        InitializeComponent();
        this.DataContext = this;
    }

    private string? _output;

    /// <summary>
    /// The settable information to the right of the label.
    /// </summary>
    public string? Output
    {
        get => _output;
        set
        {
            _output = value;
            OnPropertyChanged(nameof(Output));
        }
    }
    
    public string? Label { get; set; }
    
    public string? Tip { get; set; }
    
    public event PropertyChangedEventHandler? PropertyChanged;
    
    private void OnPropertyChanged(string propertyName)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
