using Npgsql;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GitClasesGenericas
{
    public class Connection_QueryPsg
    {
        private string ConnectionString;
        NpgsqlConnection con;

        public Connection_QueryPsg()
        {

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

        public void ExecuteQueries(string Query_)
        {
            // buscar manera de agreegar n parametros para los strored procedure
            NpgsqlCommand cmd = new NpgsqlCommand(Query_, con);
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