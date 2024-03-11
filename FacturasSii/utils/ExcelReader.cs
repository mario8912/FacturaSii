using FacturasSii.entidades;
using FacturasSii.utils;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;

namespace FacturasSii.Utils
{
    public class ExcelReader
    {
        private const string SII = "https://www2.agenciatributaria.gob.es/static_files/common/internet/dep/aplicaciones/es/aeat/ssii/fact/ws/SuministroInformacion.xsd";
        private const string SII_LR = "https://www2.agenciatributaria.gob.es/static_files/common/internet/dep/aplicaciones/es/aeat/ssii/fact/ws/SuministroLR.xsd";
        private const string SOAPENV = "http://schemas.xmlsoap.org/soap/envelope/";
        private Dictionary<int, TipoValor> _diccionarioValores;

        public void ReadExcel(string filePath)
        {
            Listas listas = new Listas();
            Excel.Application xlApp = new Excel.Application();
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
        
        public void CrearXml()
        {
            EstructuraExternaXml();
            /*foreach (var item in _diccionarioValores)
            {
                if (item.Value.Valor != null)
                {
                    Console.Write(item.Value.Campo);
                    Console.Write("\t" + item.Value.Valor);
                    Console.WriteLine();
                }
            }*/
            MessageBox.Show("Archivo creado");
        }

        
        private void EstructuraExternaXml()
        {
            XmlDocument doc = new XmlDocument();
            XmlDeclaration xmlDeclaration = doc.CreateXmlDeclaration("1.0", "UTF-8", null);
            xmlDeclaration.Encoding = "UTF-8";
            doc.AppendChild(xmlDeclaration);

            XmlElement envelope = doc.CreateElement("soapenv", "Envelope", SOAPENV);
            envelope.SetAttribute("xmlns:soapenv", SOAPENV);
            envelope.SetAttribute("xmlns:siiLR", SII_LR);
            envelope.SetAttribute("xmlns:sii", SII);
            doc.AppendChild(envelope);

            XmlElement header = doc.CreateElement("soapenv", "Header", SOAPENV);
            envelope.AppendChild(header);

            XmlElement body = doc.CreateElement("soapenv", "Body", SOAPENV);
            envelope.AppendChild(body);

            XmlElement suministroLR = doc.CreateElement("siiLR", "SuministroLRFacturasEmitidas", SII_LR);
            body.AppendChild(suministroLR);

            #region Cabecera
            suministroLR.AppendChild(CabeceraXml(doc));
            #endregion

            suministroLR.AppendChild(XmlFactura(doc));

            doc.Save(@"E:\mipc\escritorio\FacturasSii\FacturasSii\templates\nuevo.xml");
        }

        private XmlDocumentFragment XmlFactura(XmlDocument doc)
        {
            XmlElement registroLRFacturasEmitidas = doc.CreateElement("siiLR", "RegistroLRFacturasEmitidas", SII_LR);

            XmlElement periodoLiquidacion = doc.CreateElement("sii", "PeriodoLiquidacion", SII);
            registroLRFacturasEmitidas.AppendChild(periodoLiquidacion);

            XmlElement ejercicio = doc.CreateElement("sii", "Ejercicio", SII);
            ejercicio.InnerText = "2024";
            periodoLiquidacion.AppendChild(ejercicio);

            XmlElement periodo = doc.CreateElement("sii", "Periodo", SII);
            periodo.InnerText = "02"; // Febrero
            periodoLiquidacion.AppendChild(periodo);

            XmlElement idFactura = doc.CreateElement("siiLR", "IDFactura", SII_LR);
            registroLRFacturasEmitidas.AppendChild(idFactura);

            XmlElement IDEmisorFactura = doc.CreateElement("sii", "IDEmisorFactura", SII);
            idFactura.AppendChild(IDEmisorFactura);

            XmlElement nif = doc.CreateElement("sii", "NIF", SII);
            nif.InnerText = "ejemplo nif"; // NIF del cliente

            XmlElement NumSerieFacturaEmisor = doc.CreateElement("sii", "NumSerieFacturaEmisor", SII);
            idFactura.AppendChild(NumSerieFacturaEmisor);

            XmlElement FechaExpedicionFacturaEmisor = doc.CreateElement("sii", "FechaExpedicionFacturaEmisor", SII);
            idFactura.AppendChild(FechaExpedicionFacturaEmisor);

            XmlElement facturaExpeddia = doc.CreateElement("siiLR", "FacturaExpedida", SII_LR);
            registroLRFacturasEmitidas.AppendChild(facturaExpeddia);

            XmlElement TipoFactura = doc.CreateElement("sii", "TipoFactura", SII);
            facturaExpeddia.AppendChild(TipoFactura);

            XmlElement ClaveRegimenEspecialOTrascendencia = doc.CreateElement("sii", "ClaveRegimenEspecialOTrascendencia", SII);    
            facturaExpeddia.AppendChild(ClaveRegimenEspecialOTrascendencia);

            XmlElement ImporteTotal = doc.CreateElement("sii", "ImporteTotal", SII);
            facturaExpeddia.AppendChild(ImporteTotal);

            XmlElement DescripcionOperacion = doc.CreateElement("sii", "DescripcionOperacion", SII);
            facturaExpeddia.AppendChild(DescripcionOperacion);

            XmlElement Contraparte = doc.CreateElement("siiLR", "Contraparte", SII_LR);  
            facturaExpeddia.AppendChild(Contraparte);   

            XmlElement NombreRazon = doc.CreateElement("sii", "NombreRazon", SII);
            Contraparte.AppendChild(NombreRazon);

            XmlElement NIF = doc.CreateElement("sii", "NIF", SII);
            Contraparte.AppendChild(NIF);



            XmlDocumentFragment frag = doc.CreateDocumentFragment();
            frag.AppendChild(registroLRFacturasEmitidas);

            return frag;
        }

        private XmlDocumentFragment CabeceraXml(XmlDocument doc)
        {
            XmlElement cabecera = doc.CreateElement("sii", "Cabecera", SII);
            
            XmlElement idVersion = doc.CreateElement("sii", "IDVersionii", SII);
            idVersion.InnerText = "1.1";
            cabecera.AppendChild(idVersion);

            XmlElement titular = doc.CreateElement("sii", "Titular", SII);
            cabecera.AppendChild(titular);

            XmlElement nombreRazon = doc.CreateElement("sii", "NombreRazon", SII);
            nombreRazon.InnerText = "Distribuciones Rosell SL";
            titular.AppendChild(nombreRazon);

            XmlElement nif = doc.CreateElement("sii", "NIF", SII);
            nif.InnerText = "B12323648";
            titular.AppendChild(nif);

            XmlDocumentFragment frag = doc.CreateDocumentFragment();
            frag.AppendChild(cabecera);

            return frag;
        }
    }
}
