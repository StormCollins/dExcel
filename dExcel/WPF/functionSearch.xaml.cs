namespace dExcel;

using System;
using System.Collections.Generic;
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
using FuzzySharp;
using MaterialDesignColors;
using MaterialDesignThemes.Wpf;



/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class FunctionSearch : Window
{
    private List<string> matches = new List<string>();

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
        List<string> choices = new() { "d.BlackScholes", "d.Black", "d.HullWhite", "d.ModFol", "d.Prev", "d.Fol", "d.InterpContiguousArea", "d.InterpTwoColumns" };
        matches = Process.ExtractTop(SearchTerm.Text, choices).Where(x => x.Score >= 65).Select(x => x.Value).ToList();
        SearchResults.Text = "";
        if (matches.Any())
        {
            foreach (var match in matches)
            {
                SearchResults.Text += $"{match}\n";
            }
        }
    }


}
