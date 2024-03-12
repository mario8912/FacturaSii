using G = Entidades.utils.Global;
using Entidades.utils.XML;
using System.Xml;
using System.Collections.Generic;
using Entidades.utils;

namespace Datos.XML
{
    public class ConstructorXML : IConstructorXML
    {
        private static XmlElement _ultimoIndexado;
        public void EstructuraXML()
        {
            _ultimoIndexado = Envoltorio.EstructuraPrincipalXML();
        }

        public void EstructuraCabeceraXML()
        {
            _ultimoIndexado.AppendChild(Cabecera.CabeceraXml());
        }

        public void EstructuraFacturaXML(IEnumerable<Dictionary<int, TipoValor>> diccionarioValores)
        {
            foreach (Dictionary<int, TipoValor> item in diccionarioValores)
            {
                _ultimoIndexado.AppendChild(Factura.XmlFactura(item));
            }
        }

        public void GuardarXML()
        {
            G.XmlDocument.Save(@"E:\mipc\escritorio\FacturaSii\Entidades\templates\nuevo.xml");
        }
    }
}
