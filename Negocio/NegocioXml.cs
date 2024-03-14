using Datos.Excel;
using Datos.XML;
using Entidades.utils;
using System;
using System.Data;

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
            //EsructuraXML();
            LeerExcel1();
        }

        private void LeerExcel1()
        {
            ExcelReader1 excelReader = new ExcelReader1();
            DataSet dataSet = excelReader.LeerExcelRs();
            DataTable dt = excelReader.LeerExcelRs().Tables[0];

            foreach (var item in dt.AsEnumerable())
            {
                Console.WriteLine("fila" + item.ToString());   
                foreach (var item2 in item.ItemArray)
                {
                    System.Console.Write(item2.ToString() + "\t");
                }
                Console.WriteLine();
            }
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
