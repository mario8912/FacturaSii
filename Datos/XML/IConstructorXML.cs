using System.Collections.Generic;

namespace Datos.XML
{
    public interface IConstructorXML
    {
        void EstructuraXML();
        void EstructuraCabeceraXML();
        void EstructuraFacturaXML(IEnumerable<Dictionary<int, dynamic>> diccionarioValores);
        void GuardarXML();
    }
}
