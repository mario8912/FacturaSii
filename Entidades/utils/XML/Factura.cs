using System.Xml;

namespace Entidades.utils.XML
{
    internal class Factura
    {
        private const string SII = Envoltorio.SII;
        private const string SII_LR = Envoltorio.SII_LR;
        private static XmlDocument _doc = new XmlDocument();

        internal static XmlDocumentFragment XmlFactura(XmlDocument doc)
        {
            _doc = doc;

            XmlElement registroLRFacturasEmitidas = _doc.CreateElement("siiLR", "RegistroLRFacturasEmitidas", SII_LR);

            #region PeriodoLiquidacion
            registroLRFacturasEmitidas.AppendChild(XmlPeriodoLiquidacion());
            #endregion

            #region IDFactura
            registroLRFacturasEmitidas.AppendChild(XmlIDFactura());
            #endregion

            XmlElement FacturaExpedida = _doc.CreateElement("siiLR", "FacturaExpedida", SII_LR);
            registroLRFacturasEmitidas.AppendChild(FacturaExpedida);

            #region Bloque primero Factura Expedida

            XmlElement TipoFactura = _doc.CreateElement("sii", "TipoFactura", SII);
            FacturaExpedida.AppendChild(TipoFactura);

            XmlElement ClaveRegimenEspecialOTrascendencia = _doc.CreateElement("sii", "ClaveRegimenEspecialOTrascendencia", SII);
            FacturaExpedida.AppendChild(ClaveRegimenEspecialOTrascendencia);

            XmlElement ImporteTotal = _doc.CreateElement("sii", "ImporteTotal", SII);
            FacturaExpedida.AppendChild(ImporteTotal);

            XmlElement DescripcionOperacion = _doc.CreateElement("sii", "DescripcionOperacion", SII);
            FacturaExpedida.AppendChild(DescripcionOperacion);

            #endregion

            #region Contraparte
            FacturaExpedida.AppendChild(XmlContraparte());
            #endregion

            XmlElement TipoDesglose = _doc.CreateElement("sii", "TipoDesglose", SII);
            FacturaExpedida.AppendChild(TipoDesglose);

            #region DesgloseFactura
            TipoDesglose.AppendChild(XmlDesgloseFactura());
            #endregion

            XmlDocumentFragment frag = _doc.CreateDocumentFragment();
            frag.AppendChild(registroLRFacturasEmitidas);

            return frag;
        }

        private static XmlDocumentFragment XmlPeriodoLiquidacion()
        {
            XmlElement periodoLiquidacion = _doc.CreateElement("sii", "PeriodoLiquidacion", SII);

            XmlElement ejercicio = _doc.CreateElement("sii", "Ejercicio", SII);
            ejercicio.InnerText = "2024";
            periodoLiquidacion.AppendChild(ejercicio);

            XmlElement periodo = _doc.CreateElement("sii", "Periodo", SII);
            periodo.InnerText = "02"; // Febrero
            periodoLiquidacion.AppendChild(periodo);

            XmlDocumentFragment frag = _doc.CreateDocumentFragment();
            frag.AppendChild(periodoLiquidacion);

            return frag;
        }

        private static XmlDocumentFragment XmlIDFactura()
        {
            XmlElement IDFactura = _doc.CreateElement("siiLR", "IDFactura", SII_LR);

            XmlElement IDEmisorFactura = _doc.CreateElement("sii", "IDEmisorFactura", SII);
            IDFactura.AppendChild(IDEmisorFactura);

            XmlElement nif = _doc.CreateElement("sii", "NIF", SII);
            nif.InnerText = "ejemplo nif";
            IDEmisorFactura.AppendChild(nif);

            XmlElement NumSerieFacturaEmisor = _doc.CreateElement("sii", "NumSerieFacturaEmisor", SII);
            IDFactura.AppendChild(NumSerieFacturaEmisor);

            XmlElement FechaExpedicionFacturaEmisor = _doc.CreateElement("sii", "FechaExpedicionFacturaEmisor", SII);
            IDFactura.AppendChild(FechaExpedicionFacturaEmisor);

            XmlDocumentFragment frag = _doc.CreateDocumentFragment();
            frag.AppendChild(IDFactura);

            return frag;
        }   

        private static XmlDocumentFragment XmlContraparte()
        {
            XmlElement Contraparte = _doc.CreateElement("sii", "Contraparte", SII);

            XmlElement NombreRazon = _doc.CreateElement("sii", "NombreRazon", SII);
            Contraparte.AppendChild(NombreRazon);

            XmlElement NIF = _doc.CreateElement("sii", "NIF", SII); // NIF del emisor de la factura, empresa Rosell
            Contraparte.AppendChild(NIF);

            XmlDocumentFragment frag = _doc.CreateDocumentFragment();
            frag.AppendChild(Contraparte);

            return frag;
        }

        private static XmlDocumentFragment XmlDesgloseFactura()
        {
            XmlElement DesgloseFactura = _doc.CreateElement("sii", "DesgloseFactura", SII);

            XmlElement Sujeta = _doc.CreateElement("sii", "Sujeta", SII);
            DesgloseFactura.AppendChild(Sujeta);

            XmlElement NoExenta = _doc.CreateElement("sii", "NoExenta", SII);
            Sujeta.AppendChild(NoExenta);

            XmlElement TipoNoExenta = _doc.CreateElement("sii", "TipoNoExenta", SII);
            NoExenta.AppendChild(TipoNoExenta);

            XmlElement DesgloseIVA = _doc.CreateElement("sii", "DesgloseIVA", SII);
            NoExenta.AppendChild(DesgloseIVA);

            #region DetalleIVA
            DesgloseIVA.AppendChild(XmlDetalleIva());
            #endregion

            XmlDocumentFragment frag = _doc.CreateDocumentFragment();
            frag.AppendChild(DesgloseFactura);

            return frag;
        }

        private static XmlDocumentFragment XmlDetalleIva()
        {
            XmlElement DetalleIVA = _doc.CreateElement("sii", "DetalleIVA", SII);

            XmlElement TipoImpositivo = _doc.CreateElement("sii", "TipoImpositivo", SII);
            DetalleIVA.AppendChild(TipoImpositivo);

            XmlElement BaseImponible = _doc.CreateElement("sii", "BaseImponible", SII);
            DetalleIVA.AppendChild(BaseImponible);

            XmlElement CuotaRepercutida = _doc.CreateElement("sii", "CuotaRepercutida", SII);
            DetalleIVA.AppendChild(CuotaRepercutida);

            XmlDocumentFragment frag = _doc.CreateDocumentFragment();
            frag.AppendChild(DetalleIVA);

            return frag;
        }

    }
}
