using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FacturasSii.entidades
{
    internal class Cabecera
    {
        public string IDVersionSii { get; set; }
        public Titular Titular { get; set; }
        public string TipoComunicacion { get; set; }
    }

    internal class Titular
    {
        public string NombreRazon { get; set; }
        public string NIF { get; set; }
    }
}
