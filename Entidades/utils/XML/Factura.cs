using G = Entidades.utils.Global;
using System.Collections.Generic;
using System.Xml;

namespace Entidades.utils.XML
{
    public class Factura
    {
        private static Dictionary<int, TipoValor> _diccionarioValores;

        public static XmlDocumentFragment XmlFactura(Dictionary<int, TipoValor> diccionario)
        {
            _diccionarioValores = diccionario;

            XmlElement registroLRFacturasEmitidas = G.XmlDocument.CreateElement("siiLR", "RegistroLRFacturasEmitidas", G.SII_LR);

            #region PeriodoLiquidacion
            registroLRFacturasEmitidas.AppendChild(XmlPeriodoLiquidacion());
            #endregion

            #region IDFactura
            registroLRFacturasEmitidas.AppendChild(XmlIDFactura());
            #endregion

            XmlElement FacturaExpedida = G.XmlDocument.CreateElement("siiLR", "FacturaExpedida", G.SII_LR);
            registroLRFacturasEmitidas.AppendChild(FacturaExpedida);

            #region Bloque primero Factura Expedida

            XmlElement TipoFactura = G.XmlDocument.CreateElement("sii", "TipoFactura", G.SII);
            FacturaExpedida.AppendChild(TipoFactura);

            XmlElement ClaveRegimenEspecialOTrascendencia = G.XmlDocument.CreateElement("sii", "ClaveRegimenEspecialOTrascendencia", G.SII);
            FacturaExpedida.AppendChild(ClaveRegimenEspecialOTrascendencia);

            XmlElement ImporteTotal = G.XmlDocument.CreateElement("sii", "ImporteTotal", G.SII);
            FacturaExpedida.AppendChild(ImporteTotal);

            XmlElement DescripcionOperacion = G.XmlDocument.CreateElement("sii", "DescripcionOperacion", G.SII);
            FacturaExpedida.AppendChild(DescripcionOperacion);

            #endregion

            #region Contraparte
            FacturaExpedida.AppendChild(XmlContraparte());
            #endregion

            XmlElement TipoDesglose = G.XmlDocument.CreateElement("sii", "TipoDesglose", G.SII);
            FacturaExpedida.AppendChild(TipoDesglose);

            #region DesgloseFactura
            TipoDesglose.AppendChild(XmlDesgloseFactura());
            #endregion

            XmlDocumentFragment frag = G.XmlDocument.CreateDocumentFragment();
            frag.AppendChild(registroLRFacturasEmitidas);

            return frag;
        }

        private static XmlDocumentFragment XmlPeriodoLiquidacion()
        {
            XmlElement periodoLiquidacion = G.XmlDocument.CreateElement("sii", "PeriodoLiquidacion", G.SII);

            XmlElement ejercicio = G.XmlDocument.CreateElement("sii", "Ejercicio", G.SII);
            ejercicio.InnerText = FormatoDatosLista.FormatoEjercicio(_diccionarioValores[2].Valor);
            periodoLiquidacion.AppendChild(ejercicio);

            XmlElement periodo = G.XmlDocument.CreateElement("sii", "Periodo", G.SII);
            periodo.InnerText = FormatoDatosLista.FormatoPeriodo(_diccionarioValores[2].Valor); // Febrero
            periodoLiquidacion.AppendChild(periodo);

            XmlDocumentFragment frag = G.XmlDocument.CreateDocumentFragment();
            frag.AppendChild(periodoLiquidacion);

            return frag;
        }

        private static XmlDocumentFragment XmlIDFactura()
        {
            XmlElement IDFactura = G.XmlDocument.CreateElement("siiLR", "IDFactura", G.SII_LR);

            XmlElement IDEmisorFactura = G.XmlDocument.CreateElement("sii", "IDEmisorFactura", G.SII);
            IDFactura.AppendChild(IDEmisorFactura);

            XmlElement nif = G.XmlDocument.CreateElement("sii", "NIF", G.SII);
            nif.InnerText = "ejemplo nif";
            IDEmisorFactura.AppendChild(nif);

            XmlElement NumSerieFacturaEmisor = G.XmlDocument.CreateElement("sii", "NumSerieFacturaEmisor", G.SII);
            IDFactura.AppendChild(NumSerieFacturaEmisor);

            XmlElement FechaExpedicionFacturaEmisor = G.XmlDocument.CreateElement("sii", "FechaExpedicionFacturaEmisor", G.SII);
            IDFactura.AppendChild(FechaExpedicionFacturaEmisor);

            XmlDocumentFragment frag = G.XmlDocument.CreateDocumentFragment();
            frag.AppendChild(IDFactura);

            return frag;
        }   

        private static XmlDocumentFragment XmlContraparte()
        {
            XmlElement Contraparte = G.XmlDocument.CreateElement("sii", "Contraparte", G.SII);

            XmlElement NombreRazon = G.XmlDocument.CreateElement("sii", "NombreRazon", G.SII);
            Contraparte.AppendChild(NombreRazon);

            XmlElement NIF = G.XmlDocument.CreateElement("sii", "NIF", G.SII); // NIF del emisor de la factura, empresa Rosell
            Contraparte.AppendChild(NIF);

            XmlDocumentFragment frag = G.XmlDocument.CreateDocumentFragment();
            frag.AppendChild(Contraparte);

            return frag;
        }

        private static XmlDocumentFragment XmlDesgloseFactura()
        {
            XmlElement DesgloseFactura = G.XmlDocument.CreateElement("sii", "DesgloseFactura", G.SII);

            XmlElement Sujeta = G.XmlDocument.CreateElement("sii", "Sujeta", G.SII);
            DesgloseFactura.AppendChild(Sujeta);

            XmlElement NoExenta = G.XmlDocument.CreateElement("sii", "NoExenta", G.SII);
            Sujeta.AppendChild(NoExenta);

            XmlElement TipoNoExenta = G.XmlDocument.CreateElement("sii", "TipoNoExenta", G.SII);
            NoExenta.AppendChild(TipoNoExenta);

            XmlElement DesgloseIVA = G.XmlDocument.CreateElement("sii", "DesgloseIVA", G.SII);
            NoExenta.AppendChild(DesgloseIVA);

            #region DetalleIVA
            DesgloseIVA.AppendChild(XmlDetalleIva());
            #endregion

            XmlDocumentFragment frag = G.XmlDocument.CreateDocumentFragment();
            frag.AppendChild(DesgloseFactura);

            return frag;
        }

        private static XmlDocumentFragment XmlDetalleIva()
        {
            XmlElement DetalleIVA = G.XmlDocument.CreateElement("sii", "DetalleIVA", G.SII);

            XmlElement TipoImpositivo = G.XmlDocument.CreateElement("sii", "TipoImpositivo", G.SII);
            DetalleIVA.AppendChild(TipoImpositivo);

            XmlElement BaseImponible = G.XmlDocument.CreateElement("sii", "BaseImponible", G.SII);
            DetalleIVA.AppendChild(BaseImponible);

            XmlElement CuotaRepercutida = G.XmlDocument.CreateElement("sii", "CuotaRepercutida", G.SII);
            DetalleIVA.AppendChild(CuotaRepercutida);

            XmlDocumentFragment frag = G.XmlDocument.CreateDocumentFragment();
            frag.AppendChild(DetalleIVA);

            return frag;
        }

    }
}
