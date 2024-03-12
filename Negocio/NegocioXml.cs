using Datos.Excel;
using Datos.XML;

namespace Negocio
{
    public class NegocioXml
    {
        private static ConstructorXML _constructor;
        public static void CrearXml()
        {
            _constructor = new ConstructorXML();

            try
            {
                _constructor.EstructuraXML();
                _constructor.EstructuraCabeceraXML();
                CrearFacturas();
                _constructor.GuardarXML();
            }
            catch 
            {
                throw new System.Exception();
            }
        }

        public static void CrearFacturas()
        {
            _constructor.EstructuraFacturaXML(ExcelReader.LeerExcel());
        }
    }
}
