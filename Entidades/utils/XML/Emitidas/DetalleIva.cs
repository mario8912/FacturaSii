﻿using System.Xml;
using G = Entidades.utils.Global;

namespace Entidades.utils.XML.Emitidas
{       
    public class DetalleIva
    {
        public static XmlDocumentFragment XmlDetalleIva(dynamic tipoImpositivo, dynamic baseImponible, dynamic cuotaRepercutida)
        {
            XmlElement DetalleIVA = G.XmlDocument.CreateElement("sii", "DetalleIVA", G.SII);

            XmlElement TipoImpositivo = G.XmlDocument.CreateElement("sii", "TipoImpositivo", G.SII);
            //TipoImpositivo.InnerText = Helper.ReemplazarComaPunto(tipoImpositivo);
            TipoImpositivo.InnerText = Helper.ReemplazarComaPunto(tipoImpositivo) + ".00";
            DetalleIVA.AppendChild(TipoImpositivo);

            XmlElement BaseImponible = G.XmlDocument.CreateElement("sii", "BaseImponible", G.SII);
            BaseImponible.InnerText = Helper.ReemplazarComaPunto(baseImponible);
            DetalleIVA.AppendChild(BaseImponible);

            XmlElement CuotaRepercutida = G.XmlDocument.CreateElement("sii", "CuotaRepercutida", G.SII);
            CuotaRepercutida.InnerText = Helper.ReemplazarComaPunto(cuotaRepercutida);
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
