using System.Windows;

namespace dExcel.WPF
{
    /// <summary>
    /// Interaction logic for Alert.xaml
    /// </summary>
    public partial class Alert : Window
    {
        public Alert()
        {
            InitializeComponent();
        }

        private void AlertOK_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
