using Datos.Excel;
using Datos.XML;
using Entidades.utils;
using System;
using System.Data;
using System.Diagnostics;

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
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            _constructor.EstructuraXML();
            _constructor.EstructuraCabeceraXML();
            CrearFacturas();
            _constructor.GuardarXML();
            stopwatch.Stop();

            Console.WriteLine("Tiempo de ejecución: " + stopwatch.Elapsed + "ms");
        }

        private void CrearFacturas()
        {
            ExcelReader excelReader = new ExcelReader();
            _constructor.EstructuraFacturaXML(excelReader.LeerExcel());
        }
    }
}
