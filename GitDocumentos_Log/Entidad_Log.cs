using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GitClasesGenericas;

namespace GitDocumentos_Log
{
    public class Entidad_Log
    {

        public Boolean Agregar(Entidad_Log po_Clase)
        {

            string sql3 = @"add_entidad";


            Connection_QueryPsg db = new Connection_QueryPsg();
            db.OpenConection();

            db.ExecuteQueries(sql3);

            //NpgsqlConnection pgcon = new NpgsqlConnection(pgconnectionstring);
            //pgcon.Open();
            //NpgsqlCommand pgcom = new NpgsqlCommand(sql3, pgcon);
            //pgcom.CommandType = CommandType.StoredProcedure;
            //pgcom.Parameters.AddWithValue(":pEmail", "myemail@hotmail.com");
            //pgcom.Parameters.AddWithValue(":pPassword", "eikaylie78");
            //NpgsqlDataReader pgreader = pgcom.ExecuteReader();

            //while (pgreader.Read())
            //{
            //    string name = pgreader.GetString(1);
            //    string surname = pgreader.GetString(2);
            //}

            return true;
        }
    }
}
