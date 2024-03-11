using System.Xml;

namespace Entidades.utils.XML
{
    public class Envoltorio
    {
        internal const string SII = "https://www2.agenciatributaria.gob.es/static_files/common/internet/dep/aplicaciones/es/aeat/ssii/fact/ws/SuministroInformacion.xsd";
        internal const string SII_LR = "https://www2.agenciatributaria.gob.es/static_files/common/internet/dep/aplicaciones/es/aeat/ssii/fact/ws/SuministroLR.xsd";
        internal const string SOAPENV = "http://schemas.xmlsoap.org/soap/envelope/";

        public static void EstructuraExternaXml()
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
            suministroLR.AppendChild(Cabecera.CabeceraXml(doc));
            #endregion

            #region Registro de Facturas
            suministroLR.AppendChild(Factura.XmlFactura(doc));
            #endregion

            doc.Save(@"E:\mipc\escritorio\FacturaSii\templates\nuevo.xml");
        }

    }
}
