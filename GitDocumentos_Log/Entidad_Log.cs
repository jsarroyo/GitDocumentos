using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GitDocumentos;
using GitClasesGenericas;


namespace GitDocumentos_Log
{
    public class Entidad_Log
    {

        public Boolean Agregar(Entidad po_Clase, Connection_Psg po_Conexion)
        {

            string sql3 = @"add_entidad";
             
            po_Conexion.OpenConection();
            List<Argumento> elements = new List<Argumento>();
            Argumento argumento = new Argumento("id", po_Clase.Id.ToString());
            elements.Add(argumento);
            Argumento argumento2 = new Argumento("nombre", po_Clase.Nombre.ToString());
            elements.Add(argumento2);

            Argumento argumento3 = new Argumento("esvendedor", po_Clase.EsVendedor.ToString());
            elements.Add(argumento3);

            Argumento argumento4 = new Argumento("escliente", po_Clase.EsCliente.ToString());
            elements.Add(argumento4);

            Argumento argumento5 = new Argumento("escolaborador", po_Clase.EsColaborador.ToString());
            elements.Add(argumento5);

            Argumento argumento6 = new Argumento("escliente", po_Clase.EsProveedor.ToString());
            elements.Add(argumento6);

            po_Conexion.ExecuteQueries(sql3, elements);

            return true;
        }
    }
}
