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

        /// <summary>
        /// Metoda zalogowuje użytkownika do programu ComarchXL
        /// po prawidłowym zalogowaniu zwraca poprawny numer sesji
        /// </summary>
        /// <param name="SessionId"></param>
        /// <returns></returns>
        public static int  zaloguj(ref int SessionId)
        {
            var Login = new XLLoginInfo_20202
            {
                Wersja = 20202,
                ProgramID = "test",
                OpeIdent = "admin",
                OpeHaslo = "1111",
                Baza = "KABAT_TEST"
                //Baza = "import"


            };
            
            cdn_api.cdn_api.XLLogin(Login, ref SessionId);
            
          
            return SessionId;
        }

        /// <summary>
        /// Metoda wylogowuje użytkownika z programu ComarchXL
        /// </summary>
        /// <param name="SessionId"></param>
        /// <returns>0 gdy nastąpiło poprawne wylogowanie</returns>
        public static int wyloguj(ref int SessionId)
        {
            if (SessionId < 1)
            {
                MessageBox.Show("Brak otwartych sesji");
                return -1;
            }

            int retValue = cdn_api.cdn_api.XLLogout(SessionId);
            SessionId = -1;
            if (retValue != 0)
            {
                MessageBox.Show(retValue.ToString());
            }
            return retValue;
        }


        /// <summary>
        /// Funkcja dodaje nowy produkt do ComarchXL
        /// </summary>
        /// <param name="SessionId"></param>
        /// <param name="TowarID"></param>
        /// <param name="kod"></param>
        /// <param name="nazwa"></param>
        /// <param name="jm"></param>
        /// <returns></returns>
        public static int nowyProduct(int SessionId, ref int TowarID , string kod, string nazwa, string jm)
        {
            if (SessionId < 1)
            {
                MessageBox.Show("zaloguj się najpierw");
                return -1;
            }

            var Towar = new XLTowarInfo_20202
            {
                Wersja = 20202,
                Typ = 2,
                Kod = kod,
                Nazwa = nazwa,
                Jm = jm
            };

            var retValue = cdn_api.cdn_api.XLNowyTowar(SessionId, ref TowarID, Towar);

            if (retValue != 0)
            {
                switch (retValue)
                {
                    case 292:
                     MessageBox.Show($"kod błędu {retValue.ToString()} - brak jednostki miary" );
                        break;
                    case 82:
                        MessageBox.Show($"kod błędu {retValue.ToString()} - nie podano nazwy");
                        break;
                    case 83:
                        MessageBox.Show($"kod błędu {retValue.ToString()} - Jest już towar o takim kodzie");
                        break;
                    default:
                        MessageBox.Show($"kod błędu {retValue.ToString()} ");
                        break;

                }
 
            }
            return retValue;
        }


        /// <summary>
        /// Metoda tworzy nowy BOM
        /// </summary>
        /// <param name="SessionId"></param>
        /// <param name="RecepturaId"></param>
        /// <param name="kodReceptury"></param>
        /// <param name="towar"></param>
        /// <param name="ilosc"></param>
        /// <returns></returns>
        public static int nowaReceptura(int SessionId, ref int RecepturaId, string kodReceptury, string towar, string ilosc)
        {
            var RecepturaInfo = new XLRecepturaInfo_20202
            {
                Wersja = 20202,
                TypReceptury = 1,
                Symbol = kodReceptury,
                Towar = towar,
                Ilosc = ilosc
            };

            string trescBledu;

            var retValue = cdn_api.cdn_api.XLNowaReceptura(SessionId, ref RecepturaId, RecepturaInfo);

            if (retValue != 0)
            {
                switch (retValue)
                {
                    case 1:
                        trescBledu = "Błąd zapisu dokumentu";
                        break;
                    case 2:
                        trescBledu = "Błąd logoutu";
                        break;
                    case 11:
                        trescBledu = $"Nie znaleziono towaru \n towar: {towar} \n{ilosc} ";
                        break;
                    case 13:
                        trescBledu = "Nie podano ilości";
                        break;
                    case 87:
                        trescBledu = "Nie podano symbolu receptury";
                        break;
                    case 127:
                        trescBledu = $"Receptura o symbolu: {kodReceptury} już istnieje !";
                        break;
                    default:
                        trescBledu = "Nieznany błąd";
                        break;

                }
                MessageBox.Show($"kod błędu: {retValue.ToString()} treść błędu:  {trescBledu}");
            }    

            return retValue;
        }

        /// <summary>
        /// Metoda dodaje nowy składnik do BOM'u
        /// </summary>
        /// <param name="RecepturaId"></param>
        /// <param name="towar"></param>
        /// <param name="ilosc"></param>
        /// <returns></returns>
        public static int nowySkladnik(ref int RecepturaId, string towar, string ilosc)
        {

            var SkladnikRecepturyInfo = new XLSkladnikRecepturyInfo_20202
            {
                Wersja = 20202,
                Towar = towar,
                Ilosc = ilosc,
                TypPozycji = 2
            };

           // MessageBox.Show($"recepturaId {RecepturaId} towar: {towar} ilosc {ilosc} ");
            var retValue = cdn_api.cdn_api.XLDodajSkladnikReceptury(ref RecepturaId, SkladnikRecepturyInfo);
            return retValue;
        }


        /// <summary>
        /// Metoda filalizuje zapis BOM'a
        /// </summary>
        /// <param name="RecepturaId"></param>
        /// <param name="towar"></param>
        /// <param name="ilosc"></param>
        /// <returns></returns>
        public static int zamkniecieReceptury(int RecepturaId)
        {
            var ZamkniecieRecepturyInfo = new XLZamkniecieRecepturyInfo_20202
            {
                Wersja = 20202,
            };

            var retValue = cdn_api.cdn_api.XLZamknijRecepture(RecepturaId, ZamkniecieRecepturyInfo);
            return retValue;
        }

        /// <summary>
        /// Metoda Importująca bieżniki
        /// </summary>
        /// <returns></returns>
        public static int importujBiezniki(int sessionId)
        {
            
            if (sessionId <= 0)
            {
                MessageBox.Show("Brak połączenia z ComarchXL");
                return -1;

            }
            // SqlConnection jest podklasą klasy ADO.NET o nazwie Connection
            // jest przeznaczona do obsługi połączeń z bazami danych SQL Server
            SqlConnection dataConnection = new SqlConnection();
            try
            {
                // Zbudowanie ConnectionString'a za pomocą obiektu SqlConnectionStringBuilder
                SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
                builder.DataSource = "192.168.1.186";
                builder.InitialCatalog = "mwbase";
                builder.IntegratedSecurity = false;
                builder.UserID = "sa";
                builder.Password = "#slican27$x";

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
                "SELECT b.pozid, b.indeks as indeks , b.nazwa " +
                    ",CAST(coalesce(m.indeks2, '') as varchar(50)) as mieszanka_indeks " +
                    ",round(b.iloscMieszanki*0.001,3) as mieszanka_ilosc " +
                    ",CAST(coalesce(m1.indeks2, '') as varchar(50)) as mieszanka_kapa_indeks " +
                    ",round(b.iloscMieszankiKapa*0.001,3) as mieszanka_kapa_ilosc " +
                    "FROM prdkabat.biezniki as b " +
                    "left join prdkabat.mieszanki as m " +
                    "ON m.pozid = b.mieszanka " +
                    "left join prdkabat.mieszanki as m1 " +
                    "ON m1.pozid = b.mieszankaKapa " +
                    "WHERE b.pozid <= @MaxPozId " +
                    "AND b.pozid >= @MinPozId " +
                    "ORDER BY b.pozid";

                
                // parametry przekazywane do zapytania - ochrona przed sql injection
                SqlParameter param1 = new SqlParameter("@MinPozId", SqlDbType.Int, 50);
                param1.Value = 2;
                dataCommand.Parameters.Add(param1);

                SqlParameter param2 = new SqlParameter("@MaxPozId", SqlDbType.Int, 50);
                param2.Value = 160;
                dataCommand.Parameters.Add(param2);


                SqlDataReader dataReader = dataCommand.ExecuteReader();
                // obiekt SqlDataReader zawiera najbardziej aktualny wiersz pozyskiwany z bazy danych
                // metoda Read() pobiera kolejny wiersz zwraca true jeżeli został on pomyślnie odczytany
                // za pomocą GetXXX() wyodrębniamy informację z kolumny z pobranego wiersza
                // zamiast XXX wstawiamy typ danych z C# np. GetInt32(), GetString()
                // 0 oznacza pierwszą kolumnę

                string fileNameA = $@"C:\archprg\biezniki_import_102.txt";
                System.IO.File.WriteAllText(fileNameA, DateTime.Now.ToString() + "\n");

                while (dataReader.Read())
                {
                    int pozId = dataReader.GetInt32(0);
                    // obsługa wartości null
                    if (dataReader.IsDBNull(2))
                    {
                        MessageBox.Show("null");
                    }
                    else
                    {
                                            
                        string indeks = dataReader.GetString(1).Trim();
                        string nazwa = dataReader.GetString(2).Trim();
                        
                        string skladIndeks = dataReader.GetString(dataReader.GetOrdinal("mieszanka_indeks")).Trim();
                        var skladIlosc = dataReader.GetSqlDecimal(dataReader.GetOrdinal("mieszanka_ilosc"));

                        string skladIndeksK = dataReader.GetString(dataReader.GetOrdinal("mieszanka_kapa_indeks")).Trim();
                        var skladIloscK = dataReader.GetSqlDecimal(dataReader.GetOrdinal("mieszanka_kapa_ilosc"));

                        // indeks = indeks + "_01";
                        // skladIndeks = "MIE.BMA65_2";
                        // skladIndeks = "MIE.BMO65_2";
                           skladIndeks = 'X'+skladIndeks;

                        //long skladIlosc = dataReader.GetInt64(dataReader.GetOrdinal("mieszanka_ilosc"));


                        string tekst = $"\nsesja {sessionId} \n zestaw: {indeks.Trim()} \n składnik: {skladIndeks.Trim()} \n ilość składnika: {skladIlosc}";
                        

                        int SessionID = sessionId;
                        int TowarID = -1;
                        int RecepturaID = -1;
                        string jm = "szt.";
                        string kodReceptury = "podstawowa";
                        
                       int a1 = ComarchTools.nowyProduct(SessionID, ref TowarID, indeks, nazwa, jm);
                       //MessageBox.Show($"nowyProdukt = {a1}");
                       
                       int a2 = ComarchTools.nowaReceptura(SessionID, ref RecepturaID, kodReceptury, indeks, "1");
                       //MessageBox.Show($"nowaReceptura = {a2} \n receptura {RecepturaID} ");
                       
                        
                       int a3 = ComarchTools.nowySkladnik(ref RecepturaID, skladIndeks, skladIlosc.ToString());
                        //MessageBox.Show($"nowySkladnik = {a3} , skladnik = {skladIndeks} \n receptura {RecepturaID}");

                        int a4 = 0;
                        if (skladIndeksK.Length > 0)
                           a4 = ComarchTools.nowySkladnik(ref RecepturaID, skladIndeksK, skladIloscK.ToString());

                        int a5 = ComarchTools.zamkniecieReceptury(RecepturaID);
                        //MessageBox.Show($"zamkniecieReceptury = {a4}");

                        string fileNameW = $@"C:\archprg\{indeks}.txt";
                        string info = $"{tekst} ;\n" +
                            $"nowy produkt         {a1} ;\n" +
                            $"nowa receptura       {a2} ;\n" +
                            $"nowy składnik        {a3} ; {skladIndeks} ; {skladIlosc} \n" +
                            $"nowy składnik        {a4} ; {skladIndeksK} ; {skladIloscK} \n" +
                            $"zamkniecie receptury {a5}  \n";

                        System.IO.File.WriteAllText(fileNameW,info );
                                              
                        System.IO.File.AppendAllText(fileNameA, info);



                    }
                }
                // musimy zawsze zamknąć obiekt SqlDataReader
                dataReader.Close();


            }
            catch (SqlException e)
            {
                MessageBox.Show($"Błąd dostepu do bazy danych: {e.Message} \n" );
            }
            finally
            {
                // zamyka połączenie z bazą danych    
                dataConnection.Close();
            }

            return 0;
        }

        /* import drutowek */

        /// <summary>
        /// Metoda Importująca drutówki
        /// </summary>
        /// <returns></returns>
        public static int importujDrutowki(int sessionId)
        {

            if (sessionId <= 0)
            {
                MessageBox.Show("Brak połączenia z ComarchXL");
                return -1;

            }
            // SqlConnection jest podklasą klasy ADO.NET o nazwie Connection
            // jest przeznaczona do obsługi połączeń z bazami danych SQL Server
            SqlConnection dataConnection = new SqlConnection();
            try
            {
                // Zbudowanie ConnectionString'a za pomocą obiektu SqlConnectionStringBuilder
                SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
                builder.DataSource = "192.168.1.186";
                builder.InitialCatalog = "mwbase";
                builder.IntegratedSecurity = false;
                builder.UserID = "sa";
                builder.Password = "#slican27$x";

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
                "SELECT d.pozid, d.indeks , d.nazwa " +
                    ",CAST(coalesce(m.indeks2, '') as varchar(50)) as mieszanka_indeks " +
                    ",CAST(coalesce( round(d.iloscMieszanki*0.001,4)  ,0) as numeric(10,4)) as mieszanka_ilosc " +
                    ",CAST(coalesce(dr.indeks,'') as varchar(50)) as drut_indeks " +
                    ",CAST(coalesce( round(d.wagaDrutu*0.001,4)  ,0) as numeric(10,4)) as wagaDrutu " +
                    "from prdkabat.drutowki as d " +
                    "inner join prdkabat.mieszanki as m " +
                    "ON m.pozid = d.mieszanka " +
                    "inner join prdkabat.druty as dr " +
                    "ON dr.pozid = d.drut " +
                    "WHERE d.pozid <= @MaxPozId " +
                    "AND d.pozid >= @MinPozId " +
                    "ORDER BY d.pozid";


                // parametry przekazywane do zapytania - ochrona przed sql injection
                SqlParameter param1 = new SqlParameter("@MinPozId", SqlDbType.Int, 50);
                param1.Value = 1;
                dataCommand.Parameters.Add(param1);

                SqlParameter param2 = new SqlParameter("@MaxPozId", SqlDbType.Int, 50);
                param2.Value = 60;
                dataCommand.Parameters.Add(param2);


                SqlDataReader dataReader = dataCommand.ExecuteReader();
                // obiekt SqlDataReader zawiera najbardziej aktualny wiersz pozyskiwany z bazy danych
                // metoda Read() pobiera kolejny wiersz zwraca true jeżeli został on pomyślnie odczytany
                // za pomocą GetXXX() wyodrębniamy informację z kolumny z pobranego wiersza
                // zamiast XXX wstawiamy typ danych z C# np. GetInt32(), GetString()
                // 0 oznacza pierwszą kolumnę

                string fileNameA = $@"C:\archprg\drutowki_import_2.txt";
                System.IO.File.WriteAllText(fileNameA, DateTime.Now.ToString() + "\n");

                while (dataReader.Read())
                {
                    int pozId = dataReader.GetInt32(0);
                    // obsługa wartości null
                    if (dataReader.IsDBNull(2))
                    {
                        MessageBox.Show("null");
                    }
                    else
                    {

                        string indeks = dataReader.GetString(dataReader.GetOrdinal("indeks")).Trim();
                        string nazwa = dataReader.GetString(dataReader.GetOrdinal("nazwa")).Trim();

                        string skladIndeksM = dataReader.GetString(dataReader.GetOrdinal("mieszanka_indeks")).Trim();
                        var skladIloscM = dataReader.GetSqlDecimal(dataReader.GetOrdinal("mieszanka_ilosc"));

                        string skladIndeksD = dataReader.GetString(dataReader.GetOrdinal("drut_indeks")).Trim();
                        var skladIloscD = dataReader.GetSqlDecimal(dataReader.GetOrdinal("wagaDrutu"));

                        //skladIndeksM = 'X' + skladIndeksM;

                        string tekst = $"\nsesja {sessionId} \n zestaw: {indeks.Trim()} \n";


                        int SessionID = sessionId;
                        int TowarID = -1;
                        int RecepturaID = -1;
                        string jm = "szt.";
                        string kodReceptury = "podstawowa";

                        int a1 = ComarchTools.nowyProduct(SessionID, ref TowarID, indeks, nazwa, jm);
                        //MessageBox.Show($"nowyProdukt = {a1}");

                        int a2 = ComarchTools.nowaReceptura(SessionID, ref RecepturaID, kodReceptury, indeks, "1");
                        //MessageBox.Show($"nowaReceptura = {a2} \n receptura {RecepturaID} ");


                        int a3 = ComarchTools.nowySkladnik(ref RecepturaID, skladIndeksM, skladIloscM.ToString());
                        //MessageBox.Show($"nowySkladnik = {a3} , skladnik = {skladIndeks} \n receptura {RecepturaID}");

                        int a4 = 0;
                        if (skladIndeksD.Length > 0)
                            a4 = ComarchTools.nowySkladnik(ref RecepturaID, skladIndeksD, skladIloscD.ToString());

                        int a5 = ComarchTools.zamkniecieReceptury(RecepturaID);
                        //MessageBox.Show($"zamkniecieReceptury = {a4}");

                        string fileNameW = $@"C:\archprg\{indeks}.txt";
                        string info = $"{tekst} ;\n" +
                            $"nowy produkt         {a1} ;\n" +
                            $"nowa receptura       {a2} ;\n" +
                            $"nowy składnik M      {a3} ; {skladIndeksM} ; {skladIloscM} \n" +
                            $"nowy składnik D      {a4} ; {skladIndeksD} ; {skladIloscD} \n" +
                            $"zamkniecie receptury {a5}  \n";

                        System.IO.File.WriteAllText(fileNameW, info);

                        System.IO.File.AppendAllText(fileNameA, info);



                    }
                }
                // musimy zawsze zamknąć obiekt SqlDataReader
                dataReader.Close();


            }
            catch (SqlException e)
            {
                MessageBox.Show($"Błąd dostepu do bazy danych: {e.Message} \n");
            }
            finally
            {
                // zamyka połączenie z bazą danych    
                dataConnection.Close();
            }

            return 0;
        }

        /* import Kap */
        /// <summary>
        /// Metoda Importująca kapy
        /// </summary>
        /// <returns></returns>
        public static int importujKapy(int sessionId)
        {

            if (sessionId <= 0)
            {
                MessageBox.Show("Brak połączenia z ComarchXL");
                return -1;

            }
            // SqlConnection jest podklasą klasy ADO.NET o nazwie Connection
            // jest przeznaczona do obsługi połączeń z bazami danych SQL Server
            SqlConnection dataConnection = new SqlConnection();
            try
            {
                // Zbudowanie ConnectionString'a za pomocą obiektu SqlConnectionStringBuilder
                SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
                builder.DataSource = "192.168.1.186";
                builder.InitialCatalog = "mwbase";
                builder.IntegratedSecurity = false;
                builder.UserID = "sa";
                builder.Password = "#slican27$x";

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
                "SELECT k.pozid, k.indeks , k.nazwa " +
                    ",CAST(coalesce(m.indeks2, '') as varchar(50)) as mieszanka_indeks " +
                    ",CAST(coalesce( round(k.iloscMieszanki*0.001,4)  ,0) as numeric(10,4)) as mieszanka_ilosc " +
                    "from prdkabat.kapy as k " +
                    "inner join prdkabat.mieszanki as m " +
                    "ON m.pozid = k.mieszanka " +
                    "WHERE k.pozid <= @MaxPozId " +
                    "AND k.pozid >= @MinPozId " +
                    "ORDER BY k.pozid";


                // parametry przekazywane do zapytania - ochrona przed sql injection
                SqlParameter param1 = new SqlParameter("@MinPozId", SqlDbType.Int, 50);
                param1.Value = 1;
                dataCommand.Parameters.Add(param1);

                SqlParameter param2 = new SqlParameter("@MaxPozId", SqlDbType.Int, 50);
                param2.Value = 60;
                dataCommand.Parameters.Add(param2);


                SqlDataReader dataReader = dataCommand.ExecuteReader();
                // obiekt SqlDataReader zawiera najbardziej aktualny wiersz pozyskiwany z bazy danych
                // metoda Read() pobiera kolejny wiersz zwraca true jeżeli został on pomyślnie odczytany
                // za pomocą GetXXX() wyodrębniamy informację z kolumny z pobranego wiersza
                // zamiast XXX wstawiamy typ danych z C# np. GetInt32(), GetString()
                // 0 oznacza pierwszą kolumnę

                string fileNameA = $@"C:\archprg\drutowki_kapy_1.txt";
                System.IO.File.WriteAllText(fileNameA, DateTime.Now.ToString() + "\n");

                while (dataReader.Read())
                {
                    int pozId = dataReader.GetInt32(0);
                    // obsługa wartości null
                    if (dataReader.IsDBNull(2))
                    {
                        MessageBox.Show("null");
                    }
                    else
                    {

                        string indeks = dataReader.GetString(dataReader.GetOrdinal("indeks")).Trim();
                        string nazwa = dataReader.GetString(dataReader.GetOrdinal("nazwa")).Trim();

                        string skladIndeksM = dataReader.GetString(dataReader.GetOrdinal("mieszanka_indeks")).Trim();
                        var skladIloscM = dataReader.GetSqlDecimal(dataReader.GetOrdinal("mieszanka_ilosc"));

                        //skladIndeksM = 'X' + skladIndeksM;

                        string tekst = $"\nsesja {sessionId} \n zestaw: {indeks.Trim()} \n";


                        int SessionID = sessionId;
                        int TowarID = -1;
                        int RecepturaID = -1;
                        string jm = "szt.";
                        string kodReceptury = "podstawowa";

                        int a1 = ComarchTools.nowyProduct(SessionID, ref TowarID, indeks, nazwa, jm);
                        //MessageBox.Show($"nowyProdukt = {a1}");

                        int a2 = ComarchTools.nowaReceptura(SessionID, ref RecepturaID, kodReceptury, indeks, "1");
                        //MessageBox.Show($"nowaReceptura = {a2} \n receptura {RecepturaID} ");


                        int a3 = ComarchTools.nowySkladnik(ref RecepturaID, skladIndeksM, skladIloscM.ToString());
                        //MessageBox.Show($"nowySkladnik = {a3} , skladnik = {skladIndeks} \n receptura {RecepturaID}");


                        int a5 = ComarchTools.zamkniecieReceptury(RecepturaID);
                        //MessageBox.Show($"zamkniecieReceptury = {a4}");

                        string fileNameW = $@"C:\archprg\{indeks}.txt";
                        string info = $"{tekst} ;\n" +
                            $"nowy produkt         {a1} ;\n" +
                            $"nowa receptura       {a2} ;\n" +
                            $"nowy składnik M      {a3} ; {skladIndeksM} ; {skladIloscM} \n" +
                            $"zamkniecie receptury {a5}  \n";

                        System.IO.File.WriteAllText(fileNameW, info);

                        System.IO.File.AppendAllText(fileNameA, info);



                    }
                }
                // musimy zawsze zamknąć obiekt SqlDataReader
                dataReader.Close();


            }
            catch (SqlException e)
            {
                MessageBox.Show($"Błąd dostepu do bazy danych: {e.Message} \n");
            }
            finally
            {
                // zamyka połączenie z bazą danych    
                dataConnection.Close();
            }

            return 0;
        }

        /* importuj kordy cięte */

        /// <summary>
        /// Metoda Importująca kordy cięte
        /// </summary>
        /// <returns></returns>
        public static int importujKordyCiete (int sessionId)
        {

            if (sessionId <= 0)
            {
                MessageBox.Show("Brak połączenia z ComarchXL");
                return -1;

            }
            // SqlConnection jest podklasą klasy ADO.NET o nazwie Connection
            // jest przeznaczona do obsługi połączeń z bazami danych SQL Server
            SqlConnection dataConnection = new SqlConnection();
            try
            {
                // Zbudowanie ConnectionString'a za pomocą obiektu SqlConnectionStringBuilder
                SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
                builder.DataSource = "192.168.1.186";
                builder.InitialCatalog = "mwbase";
                builder.IntegratedSecurity = false;
                builder.UserID = "sa";
                builder.Password = "#slican27$x";

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
                "SELECT c.pozid, c.indeks , c.nazwa " +
                    ",CAST(coalesce(k.indeks,'') as varchar(50)) as kordGumowanyIndeks " +
                    ",1.0000 as kordGumowanyIlosc " +
                    "from prdkabat.kordy_gumowane_ciete as c " +
                    "inner join prdkabat.kordy_gumowane as k " +
                    "ON k.pozid = c.kordGumowany " +
                    "WHERE c.pozid <= @MaxPozId " +
                    "AND c.pozid >= @MinPozId " +
                    "ORDER BY c.pozid";


                // parametry przekazywane do zapytania - ochrona przed sql injection
                SqlParameter param1 = new SqlParameter("@MinPozId", SqlDbType.Int, 50);
                param1.Value = 108;
                dataCommand.Parameters.Add(param1);

                SqlParameter param2 = new SqlParameter("@MaxPozId", SqlDbType.Int, 50);
                param2.Value = 150;
                dataCommand.Parameters.Add(param2);


                string fileNameB = $@"C:\archprg\zapytanie.txt";
                System.IO.File.WriteAllText(fileNameB, DateTime.Now.ToString() + "\n");
                System.IO.File.AppendAllText(fileNameB, dataCommand.CommandText);

                SqlDataReader dataReader = dataCommand.ExecuteReader();
                // obiekt SqlDataReader zawiera najbardziej aktualny wiersz pozyskiwany z bazy danych
                // metoda Read() pobiera kolejny wiersz zwraca true jeżeli został on pomyślnie odczytany
                // za pomocą GetXXX() wyodrębniamy informację z kolumny z pobranego wiersza
                // zamiast XXX wstawiamy typ danych z C# np. GetInt32(), GetString()
                // 0 oznacza pierwszą kolumnę

                string fileNameA = $@"C:\archprg\kordy_gumowane_ciete_2.txt";
                System.IO.File.WriteAllText(fileNameA, DateTime.Now.ToString() + "\n");

              

                while (dataReader.Read())
                {
                    int pozId = dataReader.GetInt32(0);
                    // obsługa wartości null
                    if (dataReader.IsDBNull(2))
                    {
                        MessageBox.Show("null");
                    }
                    else
                    {

                        string indeks = dataReader.GetString(dataReader.GetOrdinal("indeks")).Trim();
                        string nazwa = dataReader.GetString(dataReader.GetOrdinal("nazwa")).Trim();

                        string skladIndeksM = dataReader.GetString(dataReader.GetOrdinal("kordGumowanyIndeks")).Trim();
                        var skladIloscM = dataReader.GetSqlDecimal(dataReader.GetOrdinal("kordGumowanyIlosc"));

                        //skladIndeksM = 'X' + skladIndeksM;

                        string tekst = $"\nsesja {sessionId} \n zestaw: {indeks.Trim()} \n";


                        int SessionID = sessionId;
                        int TowarID = -1;
                        int RecepturaID = -1;
                        string jm = "kg";
                        string kodReceptury = "podstawowa";

                        int a1 = ComarchTools.nowyProduct(SessionID, ref TowarID, indeks, nazwa, jm);
                        //MessageBox.Show($"nowyProdukt = {a1}");

                        int a2 = ComarchTools.nowaReceptura(SessionID, ref RecepturaID, kodReceptury, indeks, "1");
                        //MessageBox.Show($"nowaReceptura = {a2} \n receptura {RecepturaID} ");


                        int a3 = ComarchTools.nowySkladnik(ref RecepturaID, skladIndeksM, skladIloscM.ToString());
                        //MessageBox.Show($"nowySkladnik = {a3} , skladnik = {skladIndeks} \n receptura {RecepturaID}");


                        int a5 = ComarchTools.zamkniecieReceptury(RecepturaID);
                        //MessageBox.Show($"zamkniecieReceptury = {a4}");

                        string fileNameW = $@"C:\archprg\{indeks}.txt";
                        string info = $"{tekst} ;\n" +
                            $"nowy produkt         {a1} ;\n" +
                            $"nowa receptura       {a2} ;\n" +
                            $"nowy składnik M      {a3} ; {skladIndeksM} ; {skladIloscM} \n" +
                            $"zamkniecie receptury {a5}  \n";

                        System.IO.File.WriteAllText(fileNameW, info);

                        System.IO.File.AppendAllText(fileNameA, info);



                    }
                }
                // musimy zawsze zamknąć obiekt SqlDataReader
                dataReader.Close();


            }
            catch (SqlException e)
            {
                MessageBox.Show($"Błąd dostepu do bazy danych: {e.Message} \n");
            }
            finally
            {
                // zamyka połączenie z bazą danych    
                dataConnection.Close();
            }

            return 0;
        }

        /* importuj opony surowe */

        /// <summary>
        /// Metoda Importująca kordy cięte
        /// </summary>
        /// <returns></returns>
        public static int importujOponySurowe(int sessionId)
        {

            if (sessionId <= 0)
            {
                MessageBox.Show("Brak połączenia z ComarchXL");
                return -1;

            }
            // SqlConnection jest podklasą klasy ADO.NET o nazwie Connection
            // jest przeznaczona do obsługi połączeń z bazami danych SQL Server
            SqlConnection dataConnection = new SqlConnection();
            try
            {
                // Zbudowanie ConnectionString'a za pomocą obiektu SqlConnectionStringBuilder
                SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
                builder.DataSource = "192.168.1.186";
                builder.InitialCatalog = "mwbase";
                builder.IntegratedSecurity = false;
                builder.UserID = "sa";
                builder.Password = "#slican27$x";

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
                    "with kordy_ciete AS " +
                    "( " +
                    "SELECT okc.pozid " +
                    ",CAST(coalesce(c.indeks,'') as varchar(50)) as kordCietyIndeks " +
                    ",okc.opona " +
                    ",okc.warstwa " +
                    ",okc.ilosc " +
                    ",CAST(CASE okc.warstwa " +
                    "when 0 then(((o.srednicaBebnaKonf + 2 + 16) * c.szerokosc) / 1000000) * okc.ilosc * k.wagaCel " +
                    "ELSE(((o.srednicaBebnaKonf + 2 + ((okc.warstwa - 1) * 8)) * c.szerokosc) / 1000000) * okc.ilosc * k.wagaCel " +
                    "END as numeric(10, 5)) " +
                    "AS ile_gram_na_opone " +
                    "from prdkabat.opony_kordy_gumowane_ciete as okc " +
                    "inner join prdkabat.kordy_gumowane_ciete as c " +
                    "ON c.pozid = okc.kordCiety " +
                    "inner join prdkabat.opony as o " +
                    "ON o.pozid = okc.opona " +
                    "inner join prdkabat.kordy_gumowane as k " +
                    "ON k.pozid = c.kordGumowany " +
                    ") " +
                    "  " +
                    "select o.pozid " +
                    ",'OSU-' + substring(o.indeks, 3, 20) as indeks " +
                    ",'Opona Surowa ' + substring(o.nazwa, 7, 100) as nazwa " +
                    ",CAST(coalesce(bc.indeks, '') as varchar(50)) as bieznik_indeks " +
                    ",1.00 as bieznik_ilosc " +
                    ",CAST(coalesce(round((bc.iloscMieszanki + bc.iloscMieszankiKapa) * 0.001, 4), 0) as numeric(10, 4)) as bieznik_waga_kg " +
                    ",CAST(coalesce(bb.indeks, '') as varchar(50)) as bok_indeks " +
                    ",CAST(iif(bb.indeks is null, 0, 2) as numeric(10, 2)) as bok_ilosc " +
                    ",CAST(coalesce(round(bb.iloscMieszanki * 0.001 * 2, 4), 0) as numeric(10, 4)) as dwa_boki_waga_kg " +
                    ",CAST(coalesce(k.indeks, '') as varchar(50)) as kapa_indeks " +
                    ",CAST(iif(k.indeks is null, 0, 1) as numeric(10, 2)) as kapa_ilosc " +
                    ",CAST(coalesce(round(k.iloscMieszanki * 0.001, 4), 0) as numeric(10, 4)) as kapa_waga_kg " +
                    ",CAST(coalesce(ap.indeks, '') as varchar(50)) as apex_indeks " +
                    ",round(o.apexIloscGram * 0.001, 4) as apex_ilosc " +
                    ",CAST(coalesce(d.indeks, '') as varchar(50)) as drutowka_indeks " +
                    ",CAST(coalesce(round((d.iloscMieszanki + d.wagaDrutu) * 2 * 0.001, 4), 0) as numeric(10, 4)) as dwie_drutowki_waga_kg " +
                    ",CAST(iif(d.indeks is null, 0, 2) as numeric(10, 2)) as drutowka_ilosc " +
                    ",CAST(coalesce(tko.indeks, '') as varchar(50)) as tkaninaOchronna_indeks " +
                    ",CAST(ROUND(o.cal * 25.4 * PI() * o.szerokoscTkanOchr * coalesce(tko.wagaCel, 0) * 2 * 0.000001 * 0.001, 4) as numeric(10, 4)) as tkaninaOchronna_Waga_Kg " +
                    ", CAST(ROUND((select SUM(ile_gram_na_opone) * 0.001  from kordy_ciete as kc where kc.opona = o.pozid ), 4) as numeric(10, 4)) as kordy_ciete_ile_kg_na_opone " +
                    ", CAST(ROUND((select coalesce(SUM(ile_gram_na_opone) * 0.001, 0)  from kordy_ciete as kc where kc.opona = o.pozid and warstwa = 0), 4) as numeric(10, 4)) as kordy_ciete_ile_kg_na_w0 " +
                    ", CAST(ROUND((select coalesce(SUM(ile_gram_na_opone) * 0.001, 0)  from kordy_ciete as kc where kc.opona = o.pozid and warstwa = 1), 4) as numeric(10, 4)) as kordy_ciete_ile_kg_na_w1 " +
                    ", CAST(ROUND((select coalesce(SUM(ile_gram_na_opone) * 0.001, 0)  from kordy_ciete as kc where kc.opona = o.pozid and warstwa = 2), 4) as numeric(10, 4)) as kordy_ciete_ile_kg_na_w2 " +
                    ", CAST(ROUND((select coalesce(SUM(ile_gram_na_opone) * 0.001, 0)  from kordy_ciete as kc where kc.opona = o.pozid and warstwa = 3), 4) as numeric(10, 4)) as kordy_ciete_ile_kg_na_w3 " +
                    ", CAST(ROUND((select coalesce(SUM(ile_gram_na_opone) * 0.001, 0)  from kordy_ciete as kc where kc.opona = o.pozid and warstwa = 4), 4) as numeric(10, 4)) as kordy_ciete_ile_kg_na_w4 " +
                    ", CAST(ROUND((select coalesce(SUM(ile_gram_na_opone) * 0.001, 0)  from kordy_ciete as kc where kc.opona = o.pozid and warstwa = 5), 4) as numeric(10, 4)) as kordy_ciete_ile_kg_na_w5 " +
                    ", CAST(ROUND((select coalesce(SUM(ile_gram_na_opone) * 0.001, 0)  from kordy_ciete as kc where kc.opona = o.pozid and warstwa = 6), 4) as numeric(10, 4)) as kordy_ciete_ile_kg_na_w6 " +
                    ", CAST(coalesce((select kordCietyIndeks  from kordy_ciete as kc where kc.opona = o.pozid and warstwa = 0), '') as varchar(20)) as kordy_ciete_indeks_w0 " +
                    ", CAST(coalesce((select kordCietyIndeks  from kordy_ciete as kc where kc.opona = o.pozid and warstwa = 1), '') as varchar(20)) as kordy_ciete_indeks_w1 " +
                    ", CAST(coalesce((select kordCietyIndeks  from kordy_ciete as kc where kc.opona = o.pozid and warstwa = 2), '') as varchar(20)) as kordy_ciete_indeks_w2 " +
                    ", CAST(coalesce((select kordCietyIndeks  from kordy_ciete as kc where kc.opona = o.pozid and warstwa = 3), '') as varchar(20)) as kordy_ciete_indeks_w3 " +
                    ", CAST(coalesce((select kordCietyIndeks  from kordy_ciete as kc where kc.opona = o.pozid and warstwa = 4), '') as varchar(20)) as kordy_ciete_indeks_w4 " +
                    ", CAST(coalesce((select kordCietyIndeks  from kordy_ciete as kc where kc.opona = o.pozid and warstwa = 5), '') as varchar(20)) as kordy_ciete_indeks_w5 " +
                    ", CAST(coalesce((select kordCietyIndeks  from kordy_ciete as kc where kc.opona = o.pozid and warstwa = 6), '') as varchar(20)) as kordy_ciete_indeks_w6 " +
                    "from prdkabat.Opony as o " +
                    "left join prdkabat.Biezniki as bc " +
                    "ON o.bieznik = bc.pozid " +
                    "left join prdkabat.bieznik_typ as bt " +
                    "ON bt.pozid = o.typBieznika " +
                    "left join prdkabat.Biezniki as bb " +
                    "ON o.bok = bb.pozid " +
                    "left join prdkabat.kapy as k " +
                    "ON o.kapa = k.pozid " +
                    "left join prdkabat.apexy as ap " +
                    "ON ap.pozid = o.apex " +
                    "left join prdkabat.drutowki as d " +
                    "ON o.drutowka = d.pozid " +
                    "left join prdkabat.kordy_gumowane as tko " +
                    "ON tko.pozid = o.tkaninaOchronna " +
                    "WHERE o.pozid <= @MaxPozId " +
                    "AND o.pozid >= @MinPozId " +
                    "OR o.indeks = 'KOIM15311580161L' " +
                    "ORDER BY o.pozid";


                // parametry przekazywane do zapytania - ochrona przed sql injection
                SqlParameter param1 = new SqlParameter("@MinPozId", SqlDbType.Int, 50);
                param1.Value = 0;
                dataCommand.Parameters.Add(param1);

                SqlParameter param2 = new SqlParameter("@MaxPozId", SqlDbType.Int, 50);
                param2.Value = 0;
                dataCommand.Parameters.Add(param2);


                string fileNameB = $@"C:\archprg\zapytanie.txt";
                System.IO.File.WriteAllText(fileNameB, DateTime.Now.ToString() + "\n");
                System.IO.File.AppendAllText(fileNameB, dataCommand.CommandText);

                SqlDataReader dataReader = dataCommand.ExecuteReader();
                // obiekt SqlDataReader zawiera najbardziej aktualny wiersz pozyskiwany z bazy danych
                // metoda Read() pobiera kolejny wiersz zwraca true jeżeli został on pomyślnie odczytany
                // za pomocą GetXXX() wyodrębniamy informację z kolumny z pobranego wiersza
                // zamiast XXX wstawiamy typ danych z C# np. GetInt32(), GetString()
                // 0 oznacza pierwszą kolumnę

                string fileNameA = $@"C:\archprg\opony_surowe_2.txt";
                System.IO.File.WriteAllText(fileNameA, DateTime.Now.ToString() + "\n");



                while (dataReader.Read())
                {
                    int pozId = dataReader.GetInt32(0);
                    // obsługa wartości null
                    if (dataReader.IsDBNull(2))
                    {
                        MessageBox.Show("null");
                    }
                    else
                    {

                        string indeks = dataReader.GetString(dataReader.GetOrdinal("indeks")).Trim();
                        string nazwa = dataReader.GetString(dataReader.GetOrdinal("nazwa")).Trim();

                        string bieznikIndeks = dataReader.GetString(dataReader.GetOrdinal("bieznik_indeks")).Trim();
                        var bieznikIlosc = dataReader.GetSqlDecimal(dataReader.GetOrdinal("bieznik_ilosc"));

                        string bokIndeks = dataReader.GetString(dataReader.GetOrdinal("bok_indeks")).Trim();
                        var bokIlosc = dataReader.GetSqlDecimal(dataReader.GetOrdinal("bok_ilosc"));

                        string kapaIndeks = dataReader.GetString(dataReader.GetOrdinal("kapa_indeks")).Trim();
                        var kapaIlosc = dataReader.GetSqlDecimal(dataReader.GetOrdinal("kapa_ilosc"));

                        string apexIndeks = dataReader.GetString(dataReader.GetOrdinal("apex_indeks")).Trim();
                        var apexIlosc = dataReader.GetSqlDecimal(dataReader.GetOrdinal("apex_ilosc"));

                        string drutowkaIndeks = dataReader.GetString(dataReader.GetOrdinal("drutowka_indeks")).Trim();
                        var drutowkaIlosc = dataReader.GetSqlDecimal(dataReader.GetOrdinal("drutowka_ilosc"));

                        string tkaninaOchronnaIndeks = dataReader.GetString(dataReader.GetOrdinal("tkaninaOchronna_indeks")).Trim();
                        var tkaninaOchronnaIlosc = dataReader.GetSqlDecimal(dataReader.GetOrdinal("tkaninaOchronna_waga_kg"));

                        string kordyCieteW0Indeks = dataReader.GetString(dataReader.GetOrdinal("kordy_ciete_indeks_w0")).Trim();
                        var kordyCieteW0Ilosc = dataReader.GetSqlDecimal(dataReader.GetOrdinal("kordy_ciete_ile_kg_na_w0"));

                        string kordyCieteW1Indeks = dataReader.GetString(dataReader.GetOrdinal("kordy_ciete_indeks_w1")).Trim();
                        var kordyCieteW1Ilosc = dataReader.GetSqlDecimal(dataReader.GetOrdinal("kordy_ciete_ile_kg_na_w1"));

                        string kordyCieteW2Indeks = dataReader.GetString(dataReader.GetOrdinal("kordy_ciete_indeks_w2")).Trim();
                        var kordyCieteW2Ilosc = dataReader.GetSqlDecimal(dataReader.GetOrdinal("kordy_ciete_ile_kg_na_w2"));

                        string kordyCieteW3Indeks = dataReader.GetString(dataReader.GetOrdinal("kordy_ciete_indeks_w3")).Trim();
                        var kordyCieteW3Ilosc = dataReader.GetSqlDecimal(dataReader.GetOrdinal("kordy_ciete_ile_kg_na_w3"));

                        string kordyCieteW4Indeks = dataReader.GetString(dataReader.GetOrdinal("kordy_ciete_indeks_w4")).Trim();
                        var kordyCieteW4Ilosc = dataReader.GetSqlDecimal(dataReader.GetOrdinal("kordy_ciete_ile_kg_na_w4"));

                        string kordyCieteW5Indeks = dataReader.GetString(dataReader.GetOrdinal("kordy_ciete_indeks_w5")).Trim();
                        var kordyCieteW5Ilosc = dataReader.GetSqlDecimal(dataReader.GetOrdinal("kordy_ciete_ile_kg_na_w5"));

                        string kordyCieteW6Indeks = dataReader.GetString(dataReader.GetOrdinal("kordy_ciete_indeks_w6")).Trim();
                        var kordyCieteW6Ilosc = dataReader.GetSqlDecimal(dataReader.GetOrdinal("kordy_ciete_ile_kg_na_w6"));


                        string tekst = $"\nsesja {sessionId} \n zestaw: {indeks.Trim()} \n";


                        int SessionID = sessionId;
                        int TowarID = -1;
                        int RecepturaID = -1;
                        string jm = "szt.";
                        string kodReceptury = "podstawowa";
                        int produkt = -1;
                        int recepturaStart = -1;
                        int bieznik = -1;
                        int kapa = -1;
                        int bok = -1;
                        int apex = -1;
                        int drutowka = -1;
                        int tkaninaOchronna = -1;
                        int kordyCieteW0 = -1;
                        int kordyCieteW1 = -1;
                        int kordyCieteW2 = -1;
                        int kordyCieteW3 = -1;
                        int kordyCieteW4 = -1;
                        int kordyCieteW5 = -1;
                        int kordyCieteW6 = -1;
                        int recepturaStop = -1;

                        produkt = ComarchTools.nowyProduct(SessionID, ref TowarID, indeks, nazwa, jm);
                        recepturaStart = ComarchTools.nowaReceptura(SessionID, ref RecepturaID, kodReceptury, indeks, "1");

                        if (bieznikIlosc != 0)
                            bieznik = ComarchTools.nowySkladnik(ref RecepturaID, bieznikIndeks, bieznikIlosc.ToString());

                        if (bokIlosc != 0)
                            bok = ComarchTools.nowySkladnik(ref RecepturaID, bokIndeks, bokIlosc.ToString());

                        if (kapaIlosc != 0)
                            kapa = ComarchTools.nowySkladnik(ref RecepturaID, kapaIndeks, kapaIlosc.ToString());

                        if (apexIlosc != 0)
                            apex = ComarchTools.nowySkladnik(ref RecepturaID, apexIndeks, apexIlosc.ToString());

                        if (drutowkaIlosc != 0)
                            drutowka = ComarchTools.nowySkladnik(ref RecepturaID, drutowkaIndeks, drutowkaIlosc.ToString());

                        if (tkaninaOchronnaIlosc != 0)
                            tkaninaOchronna = ComarchTools.nowySkladnik(ref RecepturaID, tkaninaOchronnaIndeks, tkaninaOchronnaIlosc.ToString());

                        if (kordyCieteW0Ilosc != 0)
                            kordyCieteW0 = ComarchTools.nowySkladnik(ref RecepturaID, kordyCieteW0Indeks, kordyCieteW0Ilosc.ToString());

                        if (kordyCieteW1Ilosc != 0)
                            kordyCieteW1 = ComarchTools.nowySkladnik(ref RecepturaID, kordyCieteW1Indeks, kordyCieteW1Ilosc.ToString());

                        if (kordyCieteW2Ilosc != 0)
                            kordyCieteW2 = ComarchTools.nowySkladnik(ref RecepturaID, kordyCieteW2Indeks, kordyCieteW2Ilosc.ToString());

                        if (kordyCieteW3Ilosc != 0)
                            kordyCieteW3 = ComarchTools.nowySkladnik(ref RecepturaID, kordyCieteW3Indeks, kordyCieteW3Ilosc.ToString());

                        if (kordyCieteW4Ilosc != 0)
                            kordyCieteW4 = ComarchTools.nowySkladnik(ref RecepturaID, kordyCieteW4Indeks, kordyCieteW4Ilosc.ToString());

                        if (kordyCieteW5Ilosc != 0)
                            kordyCieteW5 = ComarchTools.nowySkladnik(ref RecepturaID, kordyCieteW5Indeks, kordyCieteW5Ilosc.ToString());

                        if (kordyCieteW6Ilosc != 0)
                            kordyCieteW6 = ComarchTools.nowySkladnik(ref RecepturaID, kordyCieteW6Indeks, kordyCieteW6Ilosc.ToString());

                        recepturaStop = ComarchTools.zamkniecieReceptury(RecepturaID);


                        string fileNameW = $@"C:\archprg\{indeks}.txt";
                        string info = $"{tekst} ;\n" +
                            $"nowy produkt         {produkt} ;\n" +
                            $"nowa receptura       {recepturaStart} ;\n" +
                            $"bieżnik              {bieznik} ; {bieznikIndeks} ; {bieznikIlosc} \n" +
                            $"bok                  {bok} ; {bokIndeks} ; {bokIlosc} \n" +
                            $"kapa                 {kapa} ; {kapaIndeks} ; {kapaIlosc} \n" +
                            $"apex                 {apex} ; {apexIndeks} ; {apexIlosc} \n" +
                            $"drutowka             {drutowka} ; {drutowkaIndeks} ; {drutowkaIlosc} \n" +
                            $"tkaninaOchronna      {tkaninaOchronna} ; {tkaninaOchronnaIndeks} ; {tkaninaOchronnaIlosc} \n" +
                            $"kordyCiete W0        {kordyCieteW0} ; {kordyCieteW0Indeks} ; {kordyCieteW0Ilosc} \n" +
                            $"kordyCiete W1        {kordyCieteW1} ; {kordyCieteW1Indeks} ; {kordyCieteW1Ilosc} \n" +
                            $"kordyCiete W2        {kordyCieteW2} ; {kordyCieteW2Indeks} ; {kordyCieteW2Ilosc} \n" +
                            $"kordyCiete W3        {kordyCieteW3} ; {kordyCieteW3Indeks} ; {kordyCieteW3Ilosc} \n" +
                            $"kordyCiete W4        {kordyCieteW4} ; {kordyCieteW4Indeks} ; {kordyCieteW4Ilosc} \n" +
                            $"kordyCiete W5        {kordyCieteW5} ; {kordyCieteW5Indeks} ; {kordyCieteW5Ilosc} \n" +
                            $"kordyCiete W6        {kordyCieteW6} ; {kordyCieteW6Indeks} ; {kordyCieteW6Ilosc} \n" +
                            $"zamkniecie receptury {recepturaStop}  \n";

                        System.IO.File.WriteAllText(fileNameW, info);

                        System.IO.File.AppendAllText(fileNameA, info);



                    }
                }
                // musimy zawsze zamknąć obiekt SqlDataReader
                dataReader.Close();
                MessageBox.Show("koniec importu opon");

            }
            catch (SqlException e)
            {
                MessageBox.Show($"Błąd dostepu do bazy danych: {e.Message} \n");
            }
            finally
            {
                // zamyka połączenie z bazą danych    
                dataConnection.Close();
            }

            return 0;
        }

        /* Importuj Opony Wulkanizowane */

        /// <summary>
        /// Metoda Importująca bieżniki
        /// </summary>
        /// <returns></returns>
        public static int importujOponyWulkanizowane(int sessionId)
        {

            if (sessionId <= 0)
            {
                MessageBox.Show("Brak połączenia z ComarchXL");
                return -1;

            }
            // SqlConnection jest podklasą klasy ADO.NET o nazwie Connection
            // jest przeznaczona do obsługi połączeń z bazami danych SQL Server
            SqlConnection dataConnection = new SqlConnection();
            try
            {
                // Zbudowanie ConnectionString'a za pomocą obiektu SqlConnectionStringBuilder
                SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
                builder.DataSource = "192.168.1.186";
                builder.InitialCatalog = "mwbase";
                builder.IntegratedSecurity = false;
                builder.UserID = "sa";
                builder.Password = "#slican27$x";

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
                "SELECT o.pozid, o.indeks , o.nazwa " +
                    "FROM prdkabat.opony as o " +
                    "WHERE o.pozid <= @MaxPozId " +
                    "AND o.pozid >= @MinPozId " +
                    "OR o.indeks = 'KOIM15311580161L' " +
                    "ORDER BY o.pozid";


                // parametry przekazywane do zapytania - ochrona przed sql injection
                SqlParameter param1 = new SqlParameter("@MinPozId", SqlDbType.Int, 50);
                param1.Value = 0;
                dataCommand.Parameters.Add(param1);

                SqlParameter param2 = new SqlParameter("@MaxPozId", SqlDbType.Int, 50);
                param2.Value = 0;
                dataCommand.Parameters.Add(param2);


                SqlDataReader dataReader = dataCommand.ExecuteReader();
                // obiekt SqlDataReader zawiera najbardziej aktualny wiersz pozyskiwany z bazy danych
                // metoda Read() pobiera kolejny wiersz zwraca true jeżeli został on pomyślnie odczytany
                // za pomocą GetXXX() wyodrębniamy informację z kolumny z pobranego wiersza
                // zamiast XXX wstawiamy typ danych z C# np. GetInt32(), GetString()
                // 0 oznacza pierwszą kolumnę

                string fileNameA = $@"C:\archprg\opony_wulkanizowane_import_102.txt";
                System.IO.File.WriteAllText(fileNameA, DateTime.Now.ToString() + "\n");

                while (dataReader.Read())
                {
                    int pozId = dataReader.GetInt32(0);
                    // obsługa wartości null
                    if (dataReader.IsDBNull(2))
                    {
                        MessageBox.Show("null");
                    }
                    else
                    {

                        string indeks = dataReader.GetString(1).Trim();
                        string nazwa = dataReader.GetString(2).Trim();

                        string skladIndeks = "OSU-"+indeks.Substring(2);
                        var skladIlosc = 1M;



                        string tekst = $"\nsesja {sessionId} \n zestaw: {indeks.Trim()} \n składnik: {skladIndeks.Trim()} \n ilość składnika: {skladIlosc}";


                        int SessionID = sessionId;
                        int TowarID = -1;
                        int RecepturaID = -1;
                        string jm = "szt.";
                        string kodReceptury = "podstawowa";

                        int a1 = ComarchTools.nowyProduct(SessionID, ref TowarID, indeks, nazwa, jm);
                        //MessageBox.Show($"nowyProdukt = {a1}");

                        int a2 = ComarchTools.nowaReceptura(SessionID, ref RecepturaID, kodReceptury, indeks, "1");
                        //MessageBox.Show($"nowaReceptura = {a2} \n receptura {RecepturaID} ");


                        int a3 = ComarchTools.nowySkladnik(ref RecepturaID, skladIndeks, skladIlosc.ToString());
                        //MessageBox.Show($"nowySkladnik = {a3} , skladnik = {skladIndeks} \n receptura {RecepturaID}");


                        int a5 = ComarchTools.zamkniecieReceptury(RecepturaID);
                        //MessageBox.Show($"zamkniecieReceptury = {a4}");

                        string fileNameW = $@"C:\archprg\{indeks}.txt";
                        string info = $"{tekst} ;\n" +
                            $"opona wulkanizowana         {a1} ;\n" +
                            $"nowa receptura       {a2} ;\n" +
                            $"opona surowa        {a3} ; {skladIndeks} ; {skladIlosc} \n" +
                            $"zamkniecie receptury {a5}  \n";

                        System.IO.File.WriteAllText(fileNameW, info);
                        System.IO.File.AppendAllText(fileNameA, info);



                    }
                }
                // musimy zawsze zamknąć obiekt SqlDataReader
                dataReader.Close();


            }
            catch (SqlException e)
            {
                MessageBox.Show($"Błąd dostepu do bazy danych: {e.Message} \n");
            }
            finally
            {
                // zamyka połączenie z bazą danych    
                dataConnection.Close();
            }

            return 0;
        }

        /* test działania API */

        public static int importTest(int sessionId)
        {

            if (sessionId <= 0)
            {
                MessageBox.Show("Brak połączenia z ComarchXL");
                return -1;

            }


            int SessionID = sessionId;
            int TowarID = -1;
            int RecepturaID = -1;
            string jm = "szt.";
            string kodReceptury = "podstawowa";
            string indeks = "T004";
            string nazwa  = "T004 - nazwa";
            string skladIndeks = "T002";
            int skladIlosc = 2;

            var a1 = ComarchTools.nowyProduct(sessionId, ref TowarID, indeks, nazwa, jm);
            MessageBox.Show($"nowyProdukt = {a1}");

            var a2 = ComarchTools.nowaReceptura(sessionId, ref RecepturaID, kodReceptury, indeks, "1");
            MessageBox.Show($"nowaReceptura = {a2} \n receptura {RecepturaID} ");

            var a3 = ComarchTools.nowySkladnik(ref RecepturaID, skladIndeks, skladIlosc.ToString());
            MessageBox.Show($"nowySkladnik = {a3} , skladnik = {skladIndeks} \n receptura {RecepturaID}");

            var a4 = ComarchTools.zamkniecieReceptury(RecepturaID);
            MessageBox.Show($"zamkniecieReceptury = {a4}");

            return 1;
        }

    }
}
