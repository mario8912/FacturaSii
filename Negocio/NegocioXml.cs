using Datos.Excel;
using Datos.XML;

namespace Negocio
{
    public class NegocioXml
    {
        private readonly IConstructorXML _constructor = new ConstructorXML();
        private readonly ExcelReader _excelReader = new ExcelReader();

        public void CrearXml()
        {
            CrearEsructuraXML();
        }

        private void CrearEsructuraXML()
        {
            _constructor.EstructuraXML()?.EstructuraCabeceraXML()?.EstructuraFacturaXML(_excelReader.GetDiccionario());
            _constructor.GuardarXML();
        }

        public void ValidarXml()
        {
            _ = new Validacion();
        }
    }
}
