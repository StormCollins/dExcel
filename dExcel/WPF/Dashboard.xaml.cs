namespace dExcel;

using System;
using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;
using MaterialDesignThemes.Wpf;

/// <summary>
/// Interaction logic for Dashboard.xaml
/// </summary>
public partial class Dashboard : Window
{
    private static Dashboard? instance = null;

    public static Dashboard Instance
    {
        get
        {
            if (instance == null)
            {
                instance = new Dashboard();
            }
            return instance;
        }
    }

    private Dashboard()
    {
        InitializeComponent();
        ShadowAssist.SetShadowDepth(this, ShadowDepth.Depth0);
        this.Icon = dExcelIcon.Source = new BitmapImage(new Uri(Directory.GetCurrentDirectory() + @"\resources\icons\dXL-logo-extra-small.ico")); 
        dExcelIcon.Source = new BitmapImage(new Uri(Directory.GetCurrentDirectory() + @"\resources\icons\dXL-logo.ico"));
        gitlabRepoLink.Source = new BitmapImage(new Uri(Directory.GetCurrentDirectory() + @"\resources\icons\follow-link-small-green.ico"));
        installationPathLink.Source = new BitmapImage(new Uri(Directory.GetCurrentDirectory() + @"\resources\icons\follow-link-small-green.ico"));
        Closing += Dashboard_Closing;
    }

    private void Dashboard_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        instance = null;
    }

    private void CloseDashboard(object sender, RoutedEventArgs e)
    {
        this.Close();
        //instance = null;
    }
}
