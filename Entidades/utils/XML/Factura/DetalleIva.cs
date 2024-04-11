using System.Xml;
using G = Entidades.utils.Global;

namespace Entidades.utils.XML.Factura
{       
    public class DetalleIva
    {
        public static XmlDocumentFragment XmlDetalleIva(dynamic tipoImpositivo, dynamic baseImponible, dynamic cuotaRepercutida)
        {
            XmlElement DetalleIVA = G.XmlDocument.CreateElement("sii", "DetalleIVA", G.SII);

            XmlElement TipoImpositivo = G.XmlDocument.CreateElement("sii", "TipoImpositivo", G.SII);
            //TipoImpositivo.InnerText = tipoImpositivo.ToString();
            TipoImpositivo.InnerText = ParseFloatTipoImp(tipoImpositivo);
            DetalleIVA.AppendChild(TipoImpositivo);

            XmlElement BaseImponible = G.XmlDocument.CreateElement("sii", "BaseImponible", G.SII);
            BaseImponible.InnerText = baseImponible.ToString();
            DetalleIVA.AppendChild(BaseImponible);

            XmlElement CuotaRepercutida = G.XmlDocument.CreateElement("sii", "CuotaRepercutida", G.SII);
            CuotaRepercutida.InnerText = cuotaRepercutida.ToString();
            DetalleIVA.AppendChild(CuotaRepercutida);

            XmlDocumentFragment frag = G.XmlDocument.CreateDocumentFragment();
            frag.AppendChild(DetalleIVA);

            return frag;
        }

        public static string ParseFloatTipoImp(dynamic tImp)
        {
            float val = float.Parse(tImp);
            return val.ToString(); 
        }
    }
}
