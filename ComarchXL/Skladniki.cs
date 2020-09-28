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

   
    public class Skladniki
    {
        public List<Skladnik> ListaSkladnikow;

        public Skladniki()
        {
            this.ListaSkladnikow = new List<Skladnik>();
            SqlConnection dataConnection = new SqlConnection();

            try
            {
                // Zbudowanie ConnectionString'a za pomocą obiektu SqlConnectionStringBuilder
                SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
                builder.DataSource = "127.0.0.1";
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
                         @"select
                            b.zestaw_matid    
                           ,b.skladn_indeks
                           ,b.skladn_nazwa
                           ,b.skladnIlosc
                           ,b.skladnJm
                           from prdkabat.view_bomy as b
                           where b.skladn_matid is not null 
                           and b.skladn_matid != 0
                           order by b.kolejn";


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

                        string skladnIndeks = dataReader.GetString(dataReader.GetOrdinal("skladn_indeks")).Trim();
                        string skladnNazwa = dataReader.GetString(dataReader.GetOrdinal("skladn_nazwa")).Trim();
                        string skladnJm = dataReader.GetString(dataReader.GetOrdinal("skladnJm")).Trim();
                        var skladnIlosc = dataReader.GetSqlDecimal(dataReader.GetOrdinal("skladnIlosc"));
                        int zestawMatId = (int)dataReader.GetSqlInt32(dataReader.GetOrdinal("zestaw_matid"));

                        this.ListaSkladnikow.Add(new Skladnik
                        {
                            zestawMatId = zestawMatId,
                            skladnIndeks = skladnIndeks,
                            skladnNazwa = skladnNazwa,
                            skladnIlosc = (decimal)skladnIlosc,
                            skladnJm = skladnJm
                        });

//                       logText = $" \n -         {skladnIndeks} ; {skladnNazwa} ; {skladnIlosc} \n";
//                       System.IO.File.AppendAllText(logFileName, logText);
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
