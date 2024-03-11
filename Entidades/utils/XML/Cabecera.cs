using System.Xml;

namespace Entidades.utils.XML
{
    internal class Cabecera
    {
        private const string SII = Envoltorio.SII;
        internal static XmlDocumentFragment CabeceraXml(XmlDocument doc)
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

            XmlElement nif = doc.CreateElement("sii", "NIF", SII); //nif del emisor encargado, en este caso a nombre de Rosell  
            nif.InnerText = "B12323648";
            titular.AppendChild(nif);

            XmlElement TipoComunicacion = doc.CreateElement("sii", "TipoComunicacion", SII);
            cabecera.AppendChild(TipoComunicacion);

            XmlDocumentFragment frag = doc.CreateDocumentFragment();
            frag.AppendChild(cabecera);

            return frag;
        }
    }
}
