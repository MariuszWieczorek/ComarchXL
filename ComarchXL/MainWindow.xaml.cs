using System.Windows;


namespace ComarchXL
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

      

        private void cmdComarchExport_Click(object sender, RoutedEventArgs e)
        {
            ComarchImport inputDialog = new ComarchImport();
            inputDialog.ShowDialog();
                
        }

        private void btnExit_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
