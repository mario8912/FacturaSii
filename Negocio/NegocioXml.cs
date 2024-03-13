using Datos.Excel;
using Datos.XML;
using Entidades.utils;

namespace Negocio
{
    public class NegocioXml
    {
        private IConstructorXML _constructor;
        private EventoProgreso _eventoProgreso;

        public void CrearXml(EventoProgreso eventoProgreso)
        {
            _eventoProgreso = eventoProgreso;
            _constructor = new ConstructorXML();
            EsructuraXML();
        }

        private void EsructuraXML()
        {
            _constructor.EstructuraXML();
            _constructor.EstructuraCabeceraXML();
            CrearFacturas();
            _constructor.GuardarXML();
        }

        private void CrearFacturas()
        {
            ExcelReader excelReader = new ExcelReader();
            _constructor.EstructuraFacturaXML(excelReader.LeerExcel(_eventoProgreso));
        }
    }
}
