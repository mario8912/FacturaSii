using Entidades.utils;
using System.Collections.Generic;

namespace Datos.XML
{
    public interface IConstructorXML
    {
        void EstructuraXML();
        void EstructuraCabeceraXML();
        void EstructuraFacturaXML(IEnumerable<Dictionary<int, TipoValor>> diccionarioValores);
        void GuardarXML();
    }
}
