using Datos.Excel;
using Datos.XML;

namespace Negocio
{
    public class NegocioXml
    {
        private static IConstructorXML _constructor;
        public static void CrearXml()
        {
            _constructor = new ConstructorXML();

            try
            {
                EsructuraXML();
            }
            catch 
            {
                throw new System.Exception();
            }
        }

        private static void EsructuraXML()
        {
            _constructor.EstructuraXML();
            _constructor.EstructuraCabeceraXML();
            CrearFacturas();
            _constructor.GuardarXML();
        }

        public static void CrearFacturas()
        {
            _constructor.EstructuraFacturaXML(ExcelReader.LeerExcel());
        }
    }
}
