using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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

namespace dExcelManager
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            this.Icon = dExcelIcon.Source = new BitmapImage(new Uri(Directory.GetCurrentDirectory() + @"\resources\icons\dXL-logo-extra-small.ico"));
            dExcelIcon.Source = new BitmapImage(new Uri(Directory.GetCurrentDirectory() + @"\resources\icons\dXL-logo.ico"));
        }

        private void Install_Click(object sender, RoutedEventArgs e)
        {
            var versionsPath = @"C:\GitLab\dExcelTools\Versions";
            if (!Directory.Exists(versionsPath))
            {
                Directory.CreateDirectory(versionsPath);
            }

        }
    }
}
