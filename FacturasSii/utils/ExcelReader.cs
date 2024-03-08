using FacturasSii.utils;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;

namespace FacturasSii.Utils
{
    public class ExcelReader
    {
        private Dictionary<int, TipoValor> _diccionarioValores;

        public void ReadExcel(string filePath)
        {
            Listas listas = new Listas();
            Application xlApp = new Excel.Application();
            Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
            _Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;

            for (int i = 2; i <= 2/*rowCount*/; i++)
            {
                _diccionarioValores = listas.DiccionarioCeldas();

                foreach (var item in _diccionarioValores)
                {
                    if (xlRange.Cells[i, item.Key] != null && xlRange.Cells[i, item.Key].Value2 != null)
                    {
                        item.Value.Valor = xlRange.Cells[i, item.Key].Value2.ToString();
                    }
                }
                //CrearXml();
                
            }

            // Cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();
            // Close and release
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            // Close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);
            // Quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            // Cleanup
            GC.Collect();

            CrearXml();
        }
        
        private void CrearXml()
        {
            EstructuraExternaXml();
            CabeceraXml();
            /*foreach (var item in _diccionarioValores)
            {
                if (item.Value.Valor != null)
                {
                    Console.Write(item.Value.Campo);
                    Console.Write("\t" + item.Value.Valor);
                    Console.WriteLine();
                }
            }*/
        }

        private void CabeceraXml()
        {
            XmlDocument doc = new XmlDocument();

            XmlElement cabecera = doc.CreateElement("sii:Cabecera");
            doc.AppendChild(cabecera);

            XmlElement idVersion = doc.CreateElement("sii", "IDVersion");
            idVersion.InnerText = "1.1";
            cabecera.AppendChild(idVersion);

            XmlElement titular = doc.CreateElement("sii", "Titular");   
            cabecera.AppendChild(titular);

            XmlElement nombreRazon = doc.CreateElement("sii", "NombreRazon");
            nombreRazon.InnerText = "Distribuciones Rosell SL";
            titular.AppendChild(nombreRazon);

            XmlElement nif = doc.CreateElement("sii", "NIF");
            nif.InnerText = "B12345678";
            titular.AppendChild(nif);

            XmlElement tipoComunicacion = doc.CreateElement("sii", "TipoComunicacion");
            tipoComunicacion.InnerText = "A0";
            cabecera.AppendChild(tipoComunicacion);

            doc.Save(@"E:\mipc\escritorio\FacturasSii\FacturasSii\templates\nuevoCabecera.xml");
        }
        private void EstructuraExternaXml()
        {
            XmlDocument doc = new XmlDocument();

            // Crear el nodo raíz (Envelope)
            XmlElement envelope = doc.CreateElement("soapenv", "Envelope", "http://schemas.xmlsoap.org/soap/envelope/");

            // Agregar los atributos de los espacios de nombres
            envelope.SetAttribute("xmlns:soapenv", "http://schemas.xmlsoap.org/soap/envelope/");
            envelope.SetAttribute("xmlns:siiLR", "https://www2.agenciatributaria.gob.es/static_files/common/internet/dep/aplicaciones/es/aeat/ssii/fact/ws/SuministroLR.xsd");
            envelope.SetAttribute("xmlns:sii", "https://www2.agenciatributaria.gob.es/static_files/common/internet/dep/aplicaciones/es/aeat/ssii/fact/ws/SuministroInformacion.xsd");

            // Crear el nodo Body dentro de Envelope
            XmlElement body = doc.CreateElement("soapenv", "Body", "http://schemas.xmlsoap.org/soap/envelope/");

            // Crear el nodo SuministroLRFacturasEmitidas dentro de Body
            XmlElement suministroLR = doc.CreateElement("siiLR", "SuministroLRFacturasEmitidas", "https://www2.agenciatributaria.gob.es/static_files/common/internet/dep/aplicaciones/es/aeat/ssii/fact/ws/SuministroLR.xsd");

            XmlElement cabecera = doc.CreateElement("sii", "Cabecera", "https://www2.agenciatributaria.gob.es/static_files/common/internet/dep/aplicaciones/es/aeat/ssii/fact/ws/SuministroLR.xsd");

            XmlElement idVersion = doc.CreateElement("sii", "IDVersionii");
            idVersion.InnerText = "1.1";
            
            XmlElement titular = doc.CreateElement("sii", "Titular");

            XmlDocumentFragment cabeceraFragment = doc.CreateDocumentFragment();
            cabeceraFragment.AppendChild(titular);

            cabecera.AppendChild(titular);

            cabecera.AppendChild(idVersion);

            suministroLR.AppendChild(cabecera);
            // Agregar SuministroLRFacturasEmitidas dentro de Body
            body.AppendChild(suministroLR);

            // Agregar Body dentro de Envelope
            envelope.AppendChild(body);

            // Agregar Envelope como nodo raíz del documento
            doc.AppendChild(envelope);

            // Mostrar el XML resultante
            doc.Save(@"E:\mipc\escritorio\FacturasSii\FacturasSii\templates\nuevo.xml");


        }
    }
}
