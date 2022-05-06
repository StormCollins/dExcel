namespace dExcel;

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;
using FuzzySharp;
using MaterialDesignColors;
using MaterialDesignThemes.Wpf;

/// <summary>
/// Interaction logic for FunctionSearch.xaml
/// </summary>
public partial class FunctionSearch : Window
{
    public string FunctionName { get; set; }

    public class FunctionMatch : INotifyPropertyChanged
    {
        private string name;

        private string category;

        private string description;

        public string Name
        {
            get => name; 
            set
            { 
                name = value;
                this.OnPropertyChanged("Name");
            }
        }

        public string Category
        {
            get => category; 
            set
            {
                category = value;
                this.OnPropertyChanged("Category");
            }
        }

        public string Description
        {
            get => description;
            set
            {
                description = value;
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

    private readonly List<(string name, string description, string category)> methods =
        RibbonController.GetExposedMethods().ToList();

    public FunctionSearch()
    {
        InitializeComponent();
        InitializeMaterialDesign();
        ShadowAssist.SetShadowDepth(this, ShadowDepth.Depth0);
        this.Icon = dExcelIcon.Source = new BitmapImage(new Uri(Directory.GetCurrentDirectory() + @"\resources\icons\dXL-logo-extra-small.ico")); 
        dExcelIcon.Source = new BitmapImage(new Uri(Directory.GetCurrentDirectory() + @"\resources\icons\dXL-logo.ico"));
    }

    private void InitializeMaterialDesign()
    {
        // Create dummy objects to force the MaterialDesign assemblies to be loaded
        // from this assembly, which causes the MaterialDesign assemblies to be searched
        // relative to this assembly's path. Otherwise, the MaterialDesign assemblies
        // are searched relative to Eclipse's path, so they're not found.
        var card = new Card();
        var hue = new Hue("Dummy", Colors.Black, Colors.White);
    }

    private void CloseDiagnostics(object sender, RoutedEventArgs e) => this.Close();

    private void SearchTerm_TextChanged(object sender, TextChangedEventArgs e)
    {
        Insert.IsEnabled = false;
        var matches =
             Process.ExtractTop(SearchTerm.Text, methods.Select(x => x.name)).Where(y => y.Score >= 65);
                
        if (matches.Any())
        {
            FunctionMatches = new();
            foreach (var match in matches)
            {
                var method = methods.First(x => x.name == match.Value);
                FunctionMatches.Add(new FunctionMatch(method));
            }
            SearchResults.ItemsSource = FunctionMatches;
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
