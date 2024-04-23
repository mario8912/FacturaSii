using System.Xml;

namespace Entidades.utils.XML
{
    public class Cabecera
    {
        public  static XmlElement UltimoIndexado { get; set; }
        public static XmlDocumentFragment CabeceraXml()
        {
            XmlElement cabecera = Global.XmlDocument.CreateElement("sii", "Cabecera", Global.SII);

            XmlElement idVersion = Global.XmlDocument.CreateElement("sii", "IDVersionSii", Global.SII);
            idVersion.InnerText = "1.1";
            cabecera.AppendChild(idVersion);

            XmlElement titular = Global.XmlDocument.CreateElement("sii", "Titular", Global.SII);
            cabecera.AppendChild(titular);

            XmlElement nombreRazon = Global.XmlDocument.CreateElement("sii", "NombreRazon", Global.SII);
            nombreRazon.InnerText = "Distribuciones Rosell SL";
            titular.AppendChild(nombreRazon);

            XmlElement nif = Global.XmlDocument.CreateElement("sii", "NIF", Global.SII);
            nif.InnerText = "B12323648";
            titular.AppendChild(nif);

            XmlElement TipoComunicacion = Global.XmlDocument.CreateElement("sii", "TipoComunicacion", Global.SII);
            TipoComunicacion.InnerText = "A0"; //A1 BA
            cabecera.AppendChild(TipoComunicacion);

            XmlDocumentFragment frag = Global.XmlDocument.CreateDocumentFragment();
            frag.AppendChild(cabecera);

            UltimoIndexado = cabecera;  

            return frag;
        }   
    }
}
