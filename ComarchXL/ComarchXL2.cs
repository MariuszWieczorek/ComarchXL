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


        public static int sprawdzCzyIstniejeIndeks(string indeks)
        {

            int ileRekordow = 0;

            SqlConnection dataConnection = new SqlConnection();
            // Zbudowanie ConnectionString'a za pomocą obiektu SqlConnectionStringBuilder
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
            builder.DataSource = "kab-svr-sql01";
            builder.InitialCatalog = "kabat";
            builder.IntegratedSecurity = false;
            builder.UserID = "rap";
            builder.Password = "rap";

            // Podstawienie zbudowanego ConnectionString'a do obiektu połączenia z bazą
            dataConnection.ConnectionString = builder.ConnectionString;
            dataConnection.Open();

            // utworzenie obiektu zawierającego treść zapytania SQL
            // polecenie tworzy obiekt SqlCommand
            SqlCommand dataCommand = new SqlCommand();

            // nadaje właściwości Connection obiektu SQLCommand wartość połączenia z bazą danych 
            dataCommand.Connection = dataConnection;
            dataCommand.CommandType = CommandType.Text;

            dataCommand.CommandText =
            @"SELECT count(*) as ile FROM [KABAT].[CDN].[TwrKarty] where Twr_Kod = @Indeks;";


            // parametry przekazywane do zapytania - ochrona przed sql injection
            SqlParameter param1 = new SqlParameter("@Indeks", SqlDbType.VarChar, 50);
            param1.Value = indeks.Trim();
            dataCommand.Parameters.Add(param1);

            SqlDataReader dataReader = dataCommand.ExecuteReader();
            // obiekt SqlDataReader zawiera najbardziej aktualny wiersz pozyskiwany z bazy danych
            // metoda Read() pobiera kolejny wiersz zwraca true jeżeli został on pomyślnie odczytany
            // za pomocą GetXXX() wyodrębniamy informację z kolumny z pobranego wiersza
            // zamiast XXX wstawiamy typ danych z C# np. GetInt32(), GetString()
            // 0 oznacza pierwszą kolumnę

            while (dataReader.Read())
            {
                int pozId = dataReader.GetInt32(0);
                // obsługa wartości null
                // Metoda sprawdza, czy value parametr jest równy DBNull.Value 
                if (dataReader.IsDBNull(0))
                {
                    MessageBox.Show("null");
                    ileRekordow = 0;
                }
                else
                {
                    ileRekordow = (int)dataReader.GetSqlInt32(dataReader.GetOrdinal("ile"));
                }
            }

            // musimy zawsze zamknąć obiekt SqlDataReader
            dataReader.Close();
            // MessageBox.Show(ileRekordow.ToString());
            return ileRekordow;
        }


        /* import Bom-ów */
        /// <summary>
        /// Metoda Importująca BOM-y nowa wersja
        /// </summary>
        /// <returns></returns>
        public static int importujBomy(int sessionId, string zestaw)
        {

            bool doXL = true;

            if (sessionId <= 0)
            {
                //MessageBox.Show("Brak połączenia z ComarchXL");
                //return -1;

            }

    

            string logFileName = $@"C:\archprg\gm\log.txt";
            string logText = "";
            System.IO.File.WriteAllText(logFileName, DateTime.Now.ToString() + "\n");


            Zestawy MojeZestawy = new Zestawy(zestaw);
            foreach (Zestaw x in MojeZestawy.ListaZestawow)
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
            string skladnIndeks;
            string zestawIndeks;

            //.Where(x => x.zestawMatId == 580)

            foreach (Zestaw x in MojeZestawy.ListaZestawow)
            {

                TowarID = -1;
                RecepturaID = -1;
                kodReceptury = "podstawowa";
                zestawIlosc = (double)x.zestawIlosc;
                zestawIndeks = x.zestawIndeks;
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

                if (zestawIndeks.StartsWith("KO"))
                {
                    zestawIndeks = zestawIndeks.Replace("-", string.Empty);
                    zestawIndeks = zestawIndeks.Trim() + "G2";
                }


                logText = $" \nzestaw:         {zestawIndeks} ; {x.zestawNazwa} ; {zestawIlosc} ; {jm} \n";
                System.IO.File.AppendAllText(logFileName, logText);


                if (doXL)
                {
                    int sprInd = sprawdzCzyIstniejeIndeks(zestawIndeks);
                    //MessageBox.Show($"{zestawIndeks} {sprInd}");
                    if (sprInd == 0)
                    {
                        int a1 = ComarchTools.nowyProduct(SessionID, ref TowarID, zestawIndeks, x.zestawNazwa, jm);
                    }
                    // MessageBox.Show($"nowyProdukt = {a1}");

                    int a2 = ComarchTools.nowaReceptura(SessionID, ref RecepturaID, kodReceptury, zestawIndeks, zestawIlosc.ToString());
                    //MessageBox.Show($"nowaReceptura = {a2} \n receptura {RecepturaID} ");
                }

                IEnumerable<Skladnik> Skladniki = MojeSkladniki.ListaSkladnikow.Where(p => p.zestawMatId == x.zestawMatId);
                foreach (Skladnik s in Skladniki)
                {

                    skladnIlosc = (double)s.skladnIlosc;
                    jm = s.skladnJm;
                    skladnIndeks = s.skladnIndeks;


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


                    if (string.Equals(s.skladnIndeks.Trim(), "MIE.OST_2"))
                    {
                        skladnIndeks = "MXX.OST_2";
                    }

                    if (string.Equals(s.skladnIndeks.Trim(), "OTT004P"))
                    {
                        skladnIndeks = "MXX.OTT004P";
                    }

                    if (string.Equals(s.skladnIndeks.Trim(), "MIE.BKR65_2"))
                    {
                        skladnIndeks = "MXX.BKR65_2";
                    }

                    if (string.Equals(s.skladnIndeks.Trim(), "MIE.BMO65_2"))
                    {
                        skladnIndeks = "MXX.BMO65_2";
                    }



                    if (doXL)
                    {
                        int a3 = ComarchTools.nowySkladnik(ref RecepturaID, skladnIndeks, skladnIlosc.ToString());
                        //MessageBox.Show($"nowySkladnik = {a3} , skladnik = {skladIndeks} \n receptura {RecepturaID}");
                    }

                    logText = $" \nskladnik:        {skladnIndeks} ; {s.skladnNazwa} ; {skladnIlosc} ; {jm} \n";
                    System.IO.File.AppendAllText(logFileName, logText);
                }

                if (doXL)
                {
                    int a5 = ComarchTools.zamkniecieReceptury(RecepturaID);
                    //MessageBox.Show($"zamkniecieReceptury = {a4}");
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

            MessageBox.Show("OK");

            return 0;


        }


    }   
}
