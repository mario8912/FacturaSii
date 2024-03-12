using G = Entidades.utils.Global;
using System.Xml;

namespace Entidades.utils.XML
{
    public class Cabecera
    {
        public  static XmlElement UltimoIndexado { get; set; }
        public static XmlDocumentFragment CabeceraXml()
        {
            XmlElement cabecera = G.XmlDocument.CreateElement("sii", "Cabecera", G.SII);

            XmlElement idVersion = G.XmlDocument.CreateElement("sii", "IDVersionii", G.SII);
            idVersion.InnerText = "1.1";
            cabecera.AppendChild(idVersion);

            XmlElement titular = G.XmlDocument.CreateElement("sii", "Titular", G.SII);
            cabecera.AppendChild(titular);

            XmlElement nombreRazon = G.XmlDocument.CreateElement("sii", "NombreRazon", G.SII);
            nombreRazon.InnerText = "Distribuciones Rosell SL";
            titular.AppendChild(nombreRazon);

            XmlElement nif = G.XmlDocument.CreateElement("sii", "NIF", G.SII);
            nif.InnerText = "B12323648";
            titular.AppendChild(nif);

            XmlElement TipoComunicacion = G.XmlDocument.CreateElement("sii", "TipoComunicacion", G.SII);
            TipoComunicacion.InnerText = "A0";
            cabecera.AppendChild(TipoComunicacion);

            XmlDocumentFragment frag = G.XmlDocument.CreateDocumentFragment();
            frag.AppendChild(cabecera);

            UltimoIndexado = cabecera;  

            return frag;
        }
    }
}
