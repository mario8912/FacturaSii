using G = Entidades.utils.Global;
using System.Collections.Generic;
using System.Xml;

namespace Entidades.utils.XML
{
    public class Envoltorio
    {
        public static XmlElement SuministroLR { get; private set; }

        public Envoltorio EstructuraPrincipalXML()
        {
            G.XmlDocument  = new XmlDocument();

            XmlDeclaration xmlDeclaration = G.XmlDocument.CreateXmlDeclaration("1.0", "UTF-8", null);
            G.XmlDocument.AppendChild(xmlDeclaration);

            XmlElement envelope = G.XmlDocument.CreateElement("soapenv", "Envelope", G.SOAPENV);
            envelope.SetAttribute("xmlns:soapenv", G.SOAPENV);
            envelope.SetAttribute("xmlns:siiLR", G.SII_LR);
            envelope.SetAttribute("xmlns:sii", G.SII);
            G.XmlDocument.AppendChild(envelope);

            XmlElement header = G.XmlDocument.CreateElement("soapenv", "Header", G.SOAPENV);
            envelope.AppendChild(header);

            XmlElement body = G.XmlDocument.CreateElement("soapenv", "Body", G.SOAPENV);
            envelope.AppendChild(body);

            SuministroLR = G.XmlDocument.CreateElement("siiLR", "SuministroLRFacturasEmitidas", G.SII_LR);
            body.AppendChild(SuministroLR);

            return this;
        }
    }
}
