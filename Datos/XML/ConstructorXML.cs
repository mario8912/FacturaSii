using G = Entidades.utils.Global;
using Entidades.utils.XML;
using System.Xml;
using System.Collections.Generic;
using System;

namespace Datos.XML
{
    public class ConstructorXML : IConstructorXML
    {
        private XmlElement _ultimoIndexado;
        public void EstructuraXML()
        {
            _ultimoIndexado = Envoltorio.EstructuraPrincipalXML();
        }

        public void EstructuraCabeceraXML()
        {
            _ultimoIndexado.AppendChild(Cabecera.CabeceraXml());
        }

        public void EstructuraFacturaXML(IEnumerable<Dictionary<int, dynamic>> diccionarioValores)
        {
            foreach (Dictionary<int, dynamic> item in diccionarioValores)
                _ultimoIndexado.AppendChild(Factura.XmlFactura(item));
        }

        public void GuardarXML()
        {
            Console.WriteLine(G.RutaGuardarXml);
            G.XmlDocument.Save(G.RutaGuardarXml);
        }
    }
}
