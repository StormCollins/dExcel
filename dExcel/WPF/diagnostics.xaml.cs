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
using MaterialDesignColors;
using MaterialDesignThemes.Wpf;

/// <summary>
/// Interaction logic for Diagnostics.xaml
/// </summary>
public partial class Diagnostics : Window
{
    private static Diagnostics instance = null;

    public static Diagnostics Instance
    {
        get
        {
            if (instance == null)
            {
                instance = new Diagnostics();
            }
            return instance;
        }
    }

    private Diagnostics()
    {
        //Directory.SetCurrentDirectory(@"C:\GitLab\dExcelTools\dExcel\dExcel\bin\Debug\net6.0-windows");
        InitializeComponent();
        InitializeMaterialDesign();
        ShadowAssist.SetShadowDepth(this, ShadowDepth.Depth0);
        this.Icon = dExcelIcon.Source = new BitmapImage(new Uri(Directory.GetCurrentDirectory() + @"\resources\icons\dXL-logo-extra-small.ico")); 
        dExcelIcon.Source = new BitmapImage(new Uri(Directory.GetCurrentDirectory() + @"\resources\icons\dXL-logo.ico"));
        gitlabRepoLink.Source = new BitmapImage(new Uri(Directory.GetCurrentDirectory() + @"\resources\icons\follow-link-small-green.ico"));
        installationPathLink.Source = new BitmapImage(new Uri(Directory.GetCurrentDirectory() + @"\resources\icons\follow-link-small-green.ico"));
        Closing += Diagnostics_Closing;
    }

    private void Diagnostics_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        instance = null;
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

    private void CloseDiagnostics(object sender, RoutedEventArgs e)
    {
        this.Close();
        instance = null;
    }
}
