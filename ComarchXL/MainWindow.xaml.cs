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

        Int32 SessionID = -1;


        private void Button_Click(object sender, RoutedEventArgs e)
        {

            ComarchTools.zaloguj(ref SessionID);
          
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            ComarchTools.wyloguj(SessionID);
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            Int32 TowarID = -1;
            var Towar = new XLTowarInfo_20201
            {
                Wersja = 20201,
                Typ = 2,
                Kod = "PPPZZZ",
                Nazwa = "PPPZZZ",
                Jm  = "szt"
            };

            
            var res = cdn_api.cdn_api.XLNowyTowar(SessionID,ref TowarID, Towar);
            MessageBox.Show(res.ToString());

        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            Int32 IDReceptury = -1;
            var RecepturaInfo = new XLRecepturaInfo_20201
            {
                Wersja = 20201,
                TypReceptury = 1,
                Symbol = "REC",
                Ilosc = "1",
                Towar = "PPPZZZ"
            };

            var SkladnikRecepturyInfo = new XLSkladnikRecepturyInfo_20201
            {
                Wersja = 20201,
                Ilosc = "1",
                Towar = "SAP001"
            };

    
            var ZamkniecieRecepturyInfo = new XLZamkniecieRecepturyInfo_20201
            {
                Wersja = 20201,
             };


            var res = cdn_api.cdn_api.XLNowaReceptura(SessionID,ref IDReceptury, RecepturaInfo);
            var res1 = cdn_api.cdn_api.XLDodajSkladnikReceptury(ref IDReceptury, SkladnikRecepturyInfo);
            var res2 = cdn_api.cdn_api.XLZamknijRecepture(IDReceptury, ZamkniecieRecepturyInfo);



        }
    }
}
