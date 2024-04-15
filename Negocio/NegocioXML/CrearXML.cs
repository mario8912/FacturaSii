using Datos.Excel;
using Datos.XML;

namespace Negocio
{
    public class CrearXML
    {
        private readonly IConstructorXML _constructor = new ConstructorXML();
        private ExcelReader _excelReader;

        public void CrearXml()
        {
            CrearEsructuraXML();
        }

        private void CrearEsructuraXML()
        {
            _excelReader = new ExcelReader();

            _constructor.EstructuraXML()?.EstructuraCabeceraXML()?.EstructuraFacturaXML(_excelReader.GetDiccionario());
            _constructor.GuardarXML();
        }
    }
}
