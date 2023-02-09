using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using ExcelDna.Integration;
using FuzzySharp;
using FuzzySharp.Extractor;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;

namespace dExcel.WPF;

/// <summary>
/// Interaction logic for FunctionSearch.xaml
/// </summary>
public partial class FunctionSearch : Window
{
    public string FunctionName { get; set; }

    private readonly Dictionary<string, string> _aqsTodExcelFunctionMapping = new()
    {
        // Dates
        ["FOLDAY"] = "d.Date_FolDay",
        ["MODFOLDAY"] = "d.Date_ModFolDay",
        ["PREVDAY"] = "d.Date_PrevDay",
        // Math
        ["DT_INTERP"] = "d.Math_Interpolate",
        ["DT_INTERP1"] = "d.Math_Interpolate",
        // Stats
        ["CHOL"] = "d.Stats_Cholesky",
        ["CORR"] = "d.Stats_CorrelationMatrix",
        ["RANDN"] = "d.Stats_NormalRandomNumbers",
        // Equities
        ["DT_VOLATILITY"] = "d.Equity_Volatility",
        ["BS"] = "d.Equity_BlackScholes",
        // Interest Rates
        ["BLACK"] = "d.IR_BlackForwardOptionPricer",
        ["Disc2ForwardRate"] = "d.IR_DiscountFactorsToForwardRate",
        ["INTCONVERT"] = "d.IR_ConvertInterestRate",
    };

    public class FunctionMatch : INotifyPropertyChanged
    {
        private string _name;

        private string _category;

        private string _description;

        public string Name
        {
            get => _name; 
            set
            { 
                _name = value;
                this.OnPropertyChanged("Name");
            }
        }

        public string Category
        {
            get => _category; 
            set
            {
                _category = value;
                this.OnPropertyChanged("Category");
            }
        }

        public string Description
        {
            get => _description;
            set
            {
                _description = value;
                this.OnPropertyChanged("Description");
            }
        }

        public FunctionMatch((string name, string description, string category) functionMatch)
        {
            this.Name = functionMatch.name;
            this.Category = functionMatch.category;
            this.Description = functionMatch.description;
        }

        private void OnPropertyChanged(string name)
        {
            if (this.PropertyChanged != null)
            {
                this.PropertyChanged(this, new PropertyChangedEventArgs(name));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
    }

    private ObservableCollection<FunctionMatch> FunctionMatches { get; set; }

    private readonly List<(string name, string description, string category)> _methods = RibbonController.GetExposedMethods().ToList();

    public FunctionSearch()
    {
        string? xllPath = Path.GetDirectoryName(ExcelDnaUtil.XllPath);
        InitializeComponent();
        this.Icon = dExcelIcon.Source = new BitmapImage(new Uri(xllPath + @"\resources\icons\dXL-logo-extra-small.ico")); 
        dExcelIcon.Source = new BitmapImage(new Uri(xllPath + @"\resources\icons\dExcel48.png"));
    }

    private void CloseFunctionSearch(object sender, RoutedEventArgs e) => this.Close();

    private void SearchTerm_TextChanged(object sender, TextChangedEventArgs e)
    {
        Insert.IsEnabled = false;

        if (_aqsTodExcelFunctionMapping.ContainsKey(SearchTerm.Text.ToUpper()))
        {
            (string name, string description, string category) method = _methods.First(y => string.CompareOrdinal(y.name, _aqsTodExcelFunctionMapping[SearchTerm.Text.ToUpper()]) == 0);
            FunctionMatches = new ObservableCollection<FunctionMatch> { new(method) };
            SearchResults.ItemsSource = FunctionMatches;
        }
        else
        {
            IEnumerable<ExtractedResult<string>> matches = Process.ExtractTop(SearchTerm.Text, _methods.Select(x => x.name)).Where(y => y.Score >= 65);
            ExtractedResult<string>[] extractedResults = matches as ExtractedResult<string>[] ?? matches.ToArray();
            if (extractedResults.Any())
            {
                FunctionMatches = new ObservableCollection<FunctionMatch>();
                foreach (ExtractedResult<string> match in extractedResults)
                {
                    (string name, string description, string category) method = _methods.First(x => x.name == match.Value);
                    FunctionMatches.Add(new FunctionMatch(method));
                }

                SearchResults.ItemsSource = FunctionMatches;
            }
        }
    }

    private void SearchResults_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        Insert.IsEnabled = true;
    }

    /// <summary>
    /// Event called when insert button is clicked. The function is inserted into the spreadsheet and the function wizard is opened.
    /// </summary>
    /// <param name="sender">Event sender.</param>
    /// <param name="e">Event arguments.</param>
    private void Insert_Click(object sender, RoutedEventArgs e)
    {
        this.FunctionName = ((FunctionMatch)SearchResults.SelectedItem).Name;
        this.Close();
    }

    /// <summary>
    /// Processes keyboard events on the main form.
    /// </summary>
    /// <param name="sender">Sender.</param>
    /// <param name="e">Key event args.</param>
    /// <exception cref="NotImplementedException"></exception>
    private void FunctionSearch_OnKeyDown(object sender, KeyEventArgs e)
    {
<<<<<<< HEAD
        switch (e.Key)
        {
            case Key.Escape:
                this.Close();
                break;
            case Key.Enter when this.SearchResults != null && this.SearchTerm.IsFocused:
                this.SearchResults.Focus();
                this.SearchResults.SelectedItem = this.SearchResults.Items[0];
                break;
=======
        if (e.Key == Key.Escape)
        {
            this.Close();
>>>>>>> b2d64016efef182e409aef583b7a9de91df59c4c
        }
    }
}
