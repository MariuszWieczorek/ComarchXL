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

        /* import Bom-ów */
        /// <summary>
        /// Metoda Importująca BOM-y nowa wersja
        /// </summary>
        /// <returns></returns>
        public static int importujBomy(int sessionId, string zestaw)
        {

            if (sessionId <= 0)
            {
                //MessageBox.Show("Brak połączenia z ComarchXL");
                //return -1;

            }

            string logFileName = $@"C:\archprg\gm\log.txt";
            string logText = "";
            System.IO.File.WriteAllText(logFileName, DateTime.Now.ToString() + "\n");


            Zestawy MojeZestawy = new Zestawy(zestaw);
            foreach(Zestaw x in MojeZestawy.ListaZestawow)
            {
                
            }

            Skladniki MojeSkladniki = new Skladniki();
            foreach (Skladnik x in MojeSkladniki.ListaSkladnikow)
            {

            }



            int SessionID = sessionId;
            int TowarID = -1;
            int RecepturaID = -1;
            string jm = "szt.";
            string kodReceptury = "podstawowa";
            double zestawIlosc;
            double skladnIlosc;

            foreach (Zestaw x in MojeZestawy.ListaZestawow)
            {

                TowarID = -1;
                RecepturaID = -1;
                kodReceptury = "podstawowa";
                zestawIlosc = (double)x.zestawIlosc;
                jm = x.zestawJm;    

                
                if (x.zestawJm.Trim() == "szt")
                    jm = "szt.";

                if (x.zestawJm.Trim() == "gram" && zestawIlosc == 1)
                {
                    jm = "kg";
                }

                if (x.zestawJm.Trim() == "gram" && zestawIlosc != 1)
                {
                    jm = "kg";
                    zestawIlosc = (double)x.zestawIlosc * 0.001;
                }


                logText = $" \nzestaw:         {x.zestawIndeks} ; {x.zestawNazwa} ; {zestawIlosc} ; {jm} \n";
                System.IO.File.AppendAllText(logFileName, logText);


                //int a1 = ComarchTools.nowyProduct(SessionID, ref TowarID, x.zestawIndeks, x.zestawNazwa, x.zestawJm);
                //MessageBox.Show($"nowyProdukt = {a1}");

                //int a2 = ComarchTools.nowaReceptura(SessionID, ref RecepturaID, kodReceptury, x.zestawIndeks, x.zestawIlosc.ToString());
                //MessageBox.Show($"nowaReceptura = {a2} \n receptura {RecepturaID} ");


                IEnumerable<Skladnik> Skladniki = MojeSkladniki.ListaSkladnikow.Where(p => p.zestawMatId == x.zestawMatId);
                foreach (Skladnik s in Skladniki)
                {

                    skladnIlosc = (double)s.skladnIlosc;
                    jm = s.skladnJm;

                   
                    if (s.skladnJm == "szt")
                        jm = "szt.";

                    if (x.zestawJm.Trim() != "gram" && s.skladnJm.Trim() == "gram")
                    {
                        jm = "kg";
                        skladnIlosc = (double)s.skladnIlosc * 0.001;
                    }

                    if (x.zestawJm.Trim() == "gram" && s.skladnJm.Trim() == "gram" && zestawIlosc == 1)
                    {
                        jm = "kg";
                    }


                    //  int a3 = ComarchTools.nowySkladnik(ref RecepturaID, s.skladnIndeks, s.skladnIlosc.ToString());
                    //MessageBox.Show($"nowySkladnik = {a3} , skladnik = {skladIndeks} \n receptura {RecepturaID}");

                    // int a5 = ComarchTools.zamkniecieReceptury(RecepturaID);
                    //MessageBox.Show($"zamkniecieReceptury = {a4}");

                    logText = $" \nskladnik:        {s.skladnIndeks} ; {s.skladnNazwa} ; {skladnIlosc} ; {jm} \n";
                    System.IO.File.AppendAllText(logFileName, logText);
                }
            }

            



            int[] numbers = { 5, 10, 8, 3, 6, 12 };

            //Query syntax:
            IEnumerable<int> numQuery1 =
                from num in numbers
                where num % 2 == 0
                orderby num
                select num;


            IEnumerable<int> numQuery2 = numbers.Where(num => num % 2 == 0).OrderBy(n => n);

            /*
            IEnumerable<Skladnik> skladnikX =
            from x in MojeSkladniki
            where x.zestawMatid > 100000
            select x.skladnIndeks, x.skladnNazwa, x.skladnIlosc;
                     
            */

            return 0;
            

        }

    }
}
