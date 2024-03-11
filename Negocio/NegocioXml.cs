using Entidades.utils;
using Entidades.utils.XML;
using System.Collections.Generic;

namespace Negocio
{
    internal class NegocioXml
    {
        public static void CrearXml(Dictionary<int, TipoValor> diccionario)
        {
            Envoltorio envoltorio = new Envoltorio();
            envoltorio.CrearXml(diccionario);
        }
    }
}
