using System.Windows.Controls;

namespace dExcelWpf;

using System.ComponentModel;

public partial class InformationalLabel : UserControl, INotifyPropertyChanged
{
    public InformationalLabel()
    {
        InitializeComponent();
        this.DataContext = this;
    }

    private string? _info;

    public string? Info
    {
        get => _info;
        set
        {
            _info = value;
            OnPropertyChanged(nameof(Info));
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
