using G = Entidades.utils.Global;
using Entidades.utils.XML;
using System.Xml;
using System.Collections.Generic;
using System;

namespace Datos.XML
{
    public class ConstructorXML : IConstructorXML
    {
        private XmlElement _suministroLR;
        private readonly Envoltorio _envoltorio = new Envoltorio();

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
            foreach (Dictionary<int, dynamic> item in diccionarioValores)
                _suministroLR.AppendChild(Factura.XmlFactura(item));

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
