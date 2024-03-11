using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Entidades
{
    public class FormatoDatosLista
    {
        public static string FormatoEjercicio(string fecha)
        {
            return fecha.Substring(6, 4);
        }   

        public static string FormatoPeriodo(string fecha)
        {
            return fecha.Substring(3, 2);

        }
    }
}
