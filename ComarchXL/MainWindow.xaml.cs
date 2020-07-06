using System;
using System.Collections.Generic;
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
using cdn_api;

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

        int SessionID = -1;


        private void cmdLogin_Click(object sender, RoutedEventArgs e)
        {

            ComarchTools.zaloguj(ref SessionID);
            Sesja.Content = SessionID.ToString();
        }

        private void cmdLogout_Click(object sender, RoutedEventArgs e)
        {
            ComarchTools.wyloguj(ref SessionID);
            Sesja.Content = SessionID.ToString();
        }

        
        private void cmdDodajTowar_Click(object sender, RoutedEventArgs e)
        {
            string kod = "PPP002";
            string nazwa = "PPP002";
            string jm = "szt";
            int TowarID = -1;

            ComarchTools.nowyProduct(SessionID, ref TowarID ,kod, nazwa, jm);
        }

        private void cmdDodajBOM_Click(object sender, RoutedEventArgs e)
        {

            int IDReceptury = -1;
            string kodReceptury = "recepturaX";
            string towarKod = "PPP002";
            string towarIlosc = "1";
            ComarchTools.nowaReceptura(SessionID, ref IDReceptury,kodReceptury,towarKod, towarIlosc);

            string skladnikKod = "SAP001";
            string skladnikIlosc = "1";
            ComarchTools.nowySkladnik(ref IDReceptury, skladnikKod, skladnikIlosc);

            string skladnikKod2 = "SAP002";
            string skladnikIlosc2 = "1";
            ComarchTools.nowySkladnik(ref IDReceptury, skladnikKod2, skladnikIlosc2);

            ComarchTools.zamkniecieReceptury(IDReceptury);
        }

        private void cmdImportBieznikow_Click(object sender, RoutedEventArgs e)
        {
            ComarchTools.importujBiezniki(SessionID);
        }

        private void ImportTest_Click(object sender, RoutedEventArgs e)
        {
            ComarchTools.importTest(SessionID);
        }

        private void cmdImportDrutowek_Click(object sender, RoutedEventArgs e)
        {
            ComarchTools.importujDrutowki(SessionID);
        }

        private void cmdImportKap_Click(object sender, RoutedEventArgs e)
        {
            ComarchTools.importujKapy(SessionID);
        }

        private void cmdImportKordyCiete_Click(object sender, RoutedEventArgs e)
        {
            ComarchTools.importujKordyCiete(SessionID);
        }

        private void cmdImportOponySurowe_Click(object sender, RoutedEventArgs e)
        {
            ComarchTools.importujOponySurowe(SessionID);
        }

        private void cmdImportOponyWulkanizowane_Click(object sender, RoutedEventArgs e)
        {
            ComarchTools.importujOponyWulkanizowane(SessionID);
        }

      
    }
}
