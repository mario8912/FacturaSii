using Datos.Excel;
using Datos.XML;
using System;

namespace Negocio
{
    public class NegocioXml
    {
        private IConstructorXML _constructor;

        public void CrearXml()
        {
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
            _constructor.EstructuraFacturaXML(excelReader.LeerExcel());
        }
    }
}
