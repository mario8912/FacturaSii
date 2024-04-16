using System.Xml;
using G = Entidades.utils.Global;

namespace Entidades.utils.XML.Factura
{       
    public class DetalleIva
    {
        public static XmlDocumentFragment XmlDetalleIva(dynamic pTipoImpositivo, dynamic pBaseImponible, dynamic pCuotaSoportada)
        {
            XmlElement DetalleIVA = G.XmlDocument.CreateElement("sii", "DetalleIVA", G.SII);

            XmlElement TipoImpositivo = G.XmlDocument.CreateElement("sii", "TipoImpositivo", G.SII);
            TipoImpositivo.InnerText = Helper.ReemplazarComaPunto(pTipoImpositivo) + ".00";
            DetalleIVA.AppendChild(TipoImpositivo);

            XmlElement BaseImponible = G.XmlDocument.CreateElement("sii", "BaseImponible", G.SII);
            BaseImponible.InnerText = Helper.ReemplazarComaPunto(pBaseImponible);
            DetalleIVA.AppendChild(BaseImponible);

            XmlElement CuotaSoportada = G.XmlDocument.CreateElement("sii", "CuotaSoportada", G.SII);
            CuotaSoportada.InnerText = Helper.ReemplazarComaPunto(pCuotaSoportada);
            DetalleIVA.AppendChild(CuotaSoportada);

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
