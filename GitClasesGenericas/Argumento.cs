using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GitClasesGenericas
{
    public class Argumento
    {
        String dato;
        String tipoDato;

        public Argumento(String psDato, String psTipoDato)
        {
            dato = psDato;
            tipoDato = psTipoDato;
        }

        public string Dato { get => dato; set => dato = value; }
        public string TipoDato { get => tipoDato; set => tipoDato = value; }
    }
}
