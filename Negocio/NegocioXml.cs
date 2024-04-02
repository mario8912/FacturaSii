using Datos.Excel;
using Datos.XML;

namespace Negocio
{
    public class NegocioXml
    {
        private readonly IConstructorXML _constructor = new ConstructorXML();
        private ExcelReader _excelReader;

        public void CrearXml()
        {
            CrearEsructuraXML();
        }

        private void CrearEsructuraXML()
        {
            LeerExcel();

            _constructor.EstructuraXML()?.EstructuraCabeceraXML()?.EstructuraFacturaXML(_excelReader.GetDiccionario());
            _constructor.GuardarXML();
        }

        private void LeerExcel()
        {
            _excelReader = new ExcelReader();
        }


        public void ValidarXml()
        {
            _ = new Validacion();
        }
    }
}
