using FuzzySharp.Extractor;

namespace dExcel;

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using ExcelDna.Integration;
using FuzzySharp;

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
        ["BLACK"] = "d.IR_Black",
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

        protected void OnPropertyChanged(string name)
        {
            if (this.PropertyChanged != null)
            {
                this.PropertyChanged(this, new PropertyChangedEventArgs(name));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
    }

    public ObservableCollection<FunctionMatch> FunctionMatches { get; set; }

    private readonly List<(string name, string description, string category)> methods = RibbonController.GetExposedMethods().ToList();

    public FunctionSearch()
    {
        var xllPath = Path.GetDirectoryName(ExcelDnaUtil.XllPath);
        InitializeComponent();
        this.Icon = dExcelIcon.Source = new BitmapImage(new Uri(xllPath + @"\resources\icons\dXL-logo-extra-small.ico")); 
        dExcelIcon.Source = new BitmapImage(new Uri(xllPath + @"\resources\icons\dExcel48.png"));
    }

    private void CloseFunctionSearch(object sender, RoutedEventArgs e) => this.Close();

    private void SearchTerm_TextChanged(object sender, TextChangedEventArgs e)
    {
        Insert.IsEnabled = false;

        if (_aqsTodExcelFunctionMapping.Keys.Contains(SearchTerm.Text.ToUpper()))
        {
            (string name, string description, string category) method = methods.First(y => string.Compare(y.name, _aqsTodExcelFunctionMapping[SearchTerm.Text.ToUpper()]) == 0);
            FunctionMatches = new() { new FunctionMatch(method) };
            SearchResults.ItemsSource = FunctionMatches;
        }
        else
        {
            IEnumerable<ExtractedResult<string>> matches = Process.ExtractTop(SearchTerm.Text, methods.Select(x => x.name)).Where(y => y.Score >= 65);
            var extractedResults = matches as ExtractedResult<string>[] ?? matches.ToArray();
            if (extractedResults.Any())
            {
                FunctionMatches = new();
                foreach (var match in extractedResults)
                {
                    var method = methods.First(x => x.name == match.Value);
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

    private void SearchResults_Unselected(object sender, RoutedEventArgs e)
    {
        Insert.IsEnabled = false;
    }

    private void Insert_Click(object sender, RoutedEventArgs e)
    {
        this.FunctionName = ((FunctionMatch)SearchResults.SelectedItem).Name;
        this.Close();
    }
}
