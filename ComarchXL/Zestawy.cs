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
using System.Collections;

namespace ComarchXL
{
    public class Zestawy
    {
        public List<Zestaw> ListaZestawow;

        public Zestawy(string zestaw)
        {
            this.ListaZestawow = new List<Zestaw>();
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
                @"select b.zestaw_matid
                ,b.zestaw_indeks
                ,b.zestaw_nazwa
                ,b.zestawIlosc
                ,b.zestawJm
                from prdkabat.view_bomy as b
                where b.zestaw_grupa = @Grupa
                group by b.zestaw_matid, b.zestaw_indeks, b.zestaw_nazwa, b.zestawIlosc ,b.zestawJm
                order by b.zestaw_indeks";


                // parametry przekazywane do zapytania - ochrona przed sql injection
                SqlParameter param1 = new SqlParameter("@Grupa", SqlDbType.VarChar, 50);
                param1.Value = zestaw.Trim();
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
                    if (dataReader.IsDBNull(2))
                    {
                        MessageBox.Show("null");
                    }
                    else
                    {

                        string zestawIndeks = dataReader.GetString(dataReader.GetOrdinal("zestaw_indeks")).Trim();
                        string zestawJm = dataReader.GetString(dataReader.GetOrdinal("zestawJm")).Trim();
                        string zestawNazwa = dataReader.GetString(dataReader.GetOrdinal("zestaw_nazwa")).Trim();
                        var zestawIlosc = dataReader.GetSqlDecimal(dataReader.GetOrdinal("zestawIlosc"));
                        int zestawMatId = (int)dataReader.GetSqlInt32(dataReader.GetOrdinal("zestaw_matid"));

                        this.ListaZestawow.Add(new Zestaw 
                                                  { zestawMatId = zestawMatId,
                                                    zestawIndeks = zestawIndeks,
                                                    zestawNazwa = zestawNazwa,
                                                    zestawIlosc = (decimal)zestawIlosc,
                                                    zestawJm    = zestawJm    
                                                   }
                        ) ;
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
        }

   
    }
}
