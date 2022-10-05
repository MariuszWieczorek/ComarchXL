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
    public class ProduktyXl
    {
        public List<Product> ListaProduktow;

        public ProduktyXl()
        {
            ListaProduktow = new List<Product>();
            SqlConnection dataConnection = new SqlConnection();
            try
            {
                // Zbudowanie ConnectionString'a za pomocą obiektu SqlConnectionStringBuilder
                SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();

                builder.DataSource = @"192.168.1.182";
                builder.InitialCatalog = "kabat_test";
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


                string query2 = $@"SELECT
                  [Twr_GidNumer] as id, [Twr_Kod] as indeks
                  FROM [KABAT].[CDN].[TwrKarty]
                  where Twr_Archiwalny = 0 and Twr_Kod <> 'A-Vista'";

                dataCommand.CommandText = query2;

                // parametry przekazywane do zapytania - ochrona przed sql injection
                SqlParameter param1 = new SqlParameter("@MinPozId", SqlDbType.Int, 50);
                param1.Value = 1;
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


                int counter = 1;

                while (dataReader.Read())
                {
                    int pozId = dataReader.GetInt32(0);
                    counter++;

                    // obsługa wartości null
                    if (dataReader.IsDBNull(0))
                    {
                        MessageBox.Show("null");
                    }
                    else
                    {

                        string indeks = dataReader.GetString(1).Trim();

                        var productToAdd = new Product
                        {
                            Id = pozId,
                            Code = indeks
                        };

                        ListaProduktow.Add(productToAdd);

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
