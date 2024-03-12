using Datos.Excel;
using Datos.XML;
using System;

namespace Negocio
{
    public class NegocioXml
    {
        private IConstructorXML _constructor;
        private readonly ExcelReader _reader = new ExcelReader();

        public void CrearXml()
        {
            _constructor = new ConstructorXML();

            //try
            //{
            //    EsructuraXML();
            //}
            //catch 
            //{
            //    throw new Exception();
            //}
            EsructuraXML();
        }

        private void EsructuraXML()
        {
            _constructor.EstructuraXML();
            _constructor.EstructuraCabeceraXML();
            CrearFacturas();
            _constructor.GuardarXML();
        }

        public void CrearFacturas()
        {
            _constructor.EstructuraFacturaXML(_reader.LeerExcel());
        }
    }
}
