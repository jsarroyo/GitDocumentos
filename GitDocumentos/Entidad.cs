using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GitDocumentos
{
    public class Entidad
    {

        private String id;
        private String nombre;
        private String direccionPrincipal;
        private String telefonoPrincipal;
        private String faxPrincipal;
        private String correoElectronico;
        private Boolean esCliente;
        private Boolean esProveedor;
        private Boolean esVendedor;
        private Boolean esColaborador;
        public String Id
        {
            get { return id; }
            set { id = value; }
        }
        public String Nombre
        {
            get { return nombre; }
            set { nombre = value; }
        }
        public String CorreoElectronico
        {
            get { return correoElectronico; }
            set { correoElectronico = value; }
        }
        public Boolean EsCliente
        {
            get { return esCliente; }
            set { esCliente = value; }
        }
        public Boolean EsProveedor
        {
            get { return esProveedor; }
            set { esProveedor = value; }
        }
        public Boolean EsVendedor
        {
            get { return esVendedor; }
            set { esVendedor = value; }
        }
        public Boolean EsColaborador
        {
            get { return esColaborador; }
            set { esColaborador = value; }
        }

        public String TelefonoPrincipal
        {
            get { return telefonoPrincipal; }
            set { telefonoPrincipal = value; }
        }
        public String FaxPrincipal
        {
            get { return faxPrincipal; }
            set { faxPrincipal = value; }
        }
        public String DireccionPrincipal
        {
            get { return direccionPrincipal; }
            set { direccionPrincipal = value; }
        }

    }
}
