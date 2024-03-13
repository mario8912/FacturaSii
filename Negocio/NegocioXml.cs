using Datos.Excel;
using Datos.XML;
using System.Threading.Tasks;

namespace Negocio
{
    public class NegocioXml
    {
        private static IConstructorXML _constructor;
        static NegocioXml()
        {
            _constructor = new ConstructorXML();

            _constructor.EstructuraXML();
            _constructor.EstructuraCabeceraXML();
            CrearFacturas();
            _constructor.GuardarXML();
        }
        private static void CrearFacturas()
        {
                ExcelReader excelReader = new ExcelReader();
                _constructor.EstructuraFacturaXML(excelReader.LeerExcel());
            
        }
    }
}
