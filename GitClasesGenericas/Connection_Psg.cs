using Npgsql;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GitClasesGenericas
{
    public class Connection_Psg
    {
        private string ConnectionString;
        //"host=localhost;database=test;user=postgres"
        NpgsqlConnection con;

        public Connection_Psg(String psConnectionString)
        {
            ConnectionString = psConnectionString; 
        }
        public Connection_Psg(String psServer, String psDataBase, String psUser, String psPassword)
        {
            ConnectionString = "pghost=" + psServer +";"+ "dbName=" + psDataBase + ";" + "login=" + psUser + ";" + "pwd=" + psPassword;
        }

        public void OpenConection()
        {
            con = new NpgsqlConnection(ConnectionString);
            con.Open();
        }
        public void CloseConnection()
        {
            con.Close();
        }

        public void ExecuteQueries(string psQuery,List<Argumento> poListaParametros)
        {
            NpgsqlCommand cmd = new NpgsqlCommand(psQuery, con);            

            foreach (Argumento b in poListaParametros)
            {
                cmd.Parameters.AddWithValue(b.Dato, b.Dato);                
            }

            cmd.ExecuteNonQuery();
        }
        public NpgsqlDataReader  DataReader(string Query_)
        {
            NpgsqlCommand cmd = new NpgsqlCommand(Query_, con);
            NpgsqlDataReader dr = cmd.ExecuteReader();
            return dr;
        }
        public object ShowDataInGridView(string Query_)
        {
            NpgsqlDataAdapter  dr = new NpgsqlDataAdapter(Query_, ConnectionString);
            DataSet ds = new DataSet();
            dr.Fill(ds);
            object dataum = ds.Tables[0];

            return dataum;
        }
         
    }
}