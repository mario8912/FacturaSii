using G = Entidades.utils.Global;
using Entidades.utils.XML;
using System.Xml;
using System.Collections.Generic;
using System;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using Entidades.utils.XML.Factura;
using System.Security.Cryptography;

namespace Datos.XML
{
    public class ConstructorXML : IConstructorXML
    {
        private XmlElement _suministroLR;
        private XmlDocumentFragment _desgloseIVA;

        private readonly Envoltorio _envoltorio = new Envoltorio();
        private IEnumerable<Dictionary<int, dynamic>> _listaDiccionarioValores;

        public ConstructorXML EstructuraXML()
        {
            _envoltorio.EstructuraPrincipalXML();
            _suministroLR = Envoltorio.SuministroLR;
            return this;
        }

        public ConstructorXML EstructuraCabeceraXML()
        {
            _suministroLR.AppendChild(Cabecera.CabeceraXml());

            return this;
        }

        public void EstructuraFacturaXML(IEnumerable<Dictionary<int, dynamic>> diccionarioValores)
        {
            _listaDiccionarioValores = diccionarioValores;

            foreach (Dictionary<int, dynamic> item in diccionarioValores)
                _suministroLR.AppendChild(FacturaEmitida.XmlFactura(item));
        }

        public void GuardarXML()
        {
            try
            {
                Console.WriteLine(G.RutaGuardarXml);
                G.XmlDocument.Save(G.RutaGuardarXml);
            }
            catch (Exception)
            {
                throw new Exception("Error al guardar el archivo XML");
            }
            
        }
    }
}
