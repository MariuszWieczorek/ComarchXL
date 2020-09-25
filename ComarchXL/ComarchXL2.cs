using cdn_api;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
// przestrzeń nazw zawiera wiele typów danych używanych przez ADO.NET 
using System.Data;
// zawiera klasy dostawców danych
using System.Data.SqlClient;

namespace ComarchXL
{
    public static partial class ComarchTools
    {

        /* import Kap */
        /// <summary>
        /// Metoda Importująca kapy nowa wersja
        /// </summary>
        /// <returns></returns>
        public static int importujKapy2(int sessionId)
        {

            if (sessionId <= 0)
            {
                //MessageBox.Show("Brak połączenia z ComarchXL");
                //return -1;

            }

            string logFileName = $@"C:\archprg\gm\log.txt";
            string logText = "";
            System.IO.File.WriteAllText(logFileName, DateTime.Now.ToString() + "\n");


            Zestawy MojeZestawy = new Zestawy();
            foreach(Zestaw x in MojeZestawy.ListaZestawow)
            {
                logText = $" \nzestaw:         {x.zestawIndeks} ; {x.zestawNazwa} ; {x.zestawIlosc} \n";
                System.IO.File.AppendAllText(logFileName, logText);
                
            }

            Skladniki MojeSkladniki = new Skladniki();
            foreach (Skladnik x in MojeSkladniki.ListaSkladnikow)
            {
                logText = $" \nskladnik:        {x.skladnIndeks} ; {x.skladnNazwa} ; {x.skladnIlosc} \n";
                System.IO.File.AppendAllText(logFileName, logText);

            }

            

            return 0;
            

        }

    }
}
