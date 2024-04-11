using Entidades.utils.XML.Factura;
using System.Collections.Generic;
using System.Xml;
using G = Entidades.utils.Global;
using H = Entidades.utils.Helper;

namespace Entidades.utils.XML
{
    public class FacturaEmitida
    {
        private static Dictionary<int, dynamic> _diccionarioValores;

        public static XmlDocumentFragment XmlFactura(Dictionary<int, dynamic> diccionario)
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
            TipoFactura.InnerText = "F1";
            FacturaExpedida.AppendChild(TipoFactura);

            XmlElement ClaveRegimenEspecialOTrascendencia = G.XmlDocument.CreateElement("sii", "ClaveRegimenEspecialOTrascendencia", G.SII);
            ClaveRegimenEspecialOTrascendencia.InnerText = "01";
            FacturaExpedida.AppendChild(ClaveRegimenEspecialOTrascendencia);

            XmlElement ImporteTotal = G.XmlDocument.CreateElement("sii", "ImporteTotal", G.SII);
            ImporteTotal.InnerText = H.SumaBases(_diccionarioValores); //base1
            FacturaExpedida.AppendChild(ImporteTotal);

            XmlElement DescripcionOperacion = G.XmlDocument.CreateElement("sii", "DescripcionOperacion", G.SII);
            DescripcionOperacion.InnerText = string.Format("Venta de productos de hostelería a {0}, f. {1}", _diccionarioValores[3], _diccionarioValores[0]); //descripcion   
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
            ejercicio.InnerText = H.FormatoEjercicio(_diccionarioValores[1]); //ejercicio
            periodoLiquidacion.AppendChild(ejercicio);

            XmlElement periodo = G.XmlDocument.CreateElement("sii", "Periodo", G.SII);
            periodo.InnerText = H.FormatoPeriodo(_diccionarioValores[1]); //periodo
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
            nif.InnerText = "B12323648";
            IDEmisorFactura.AppendChild(nif);

            XmlElement NumSerieFacturaEmisor = G.XmlDocument.CreateElement("sii", "NumSerieFacturaEmisor", G.SII);
            NumSerieFacturaEmisor.InnerText = _diccionarioValores[0]; //numSerie
            IDFactura.AppendChild(NumSerieFacturaEmisor);

            XmlElement FechaExpedicionFacturaEmisor = G.XmlDocument.CreateElement("sii", "FechaExpedicionFacturaEmisor", G.SII);
            FechaExpedicionFacturaEmisor.InnerText = _diccionarioValores[1]; //fechaExpedicion
            IDFactura.AppendChild(FechaExpedicionFacturaEmisor);

            XmlDocumentFragment frag = G.XmlDocument.CreateDocumentFragment();
            frag.AppendChild(IDFactura);

            return frag;
        }

        private static XmlDocumentFragment XmlContraparte()
        {
            XmlElement Contraparte = G.XmlDocument.CreateElement("sii", "Contraparte", G.SII);

            #region NPMBRE RAZÓN
            XmlElement NombreRazon = G.XmlDocument.CreateElement("sii", "NombreRazon", G.SII);

            NombreRazon.InnerText = _diccionarioValores[3]; //nombreRazon
            Contraparte.AppendChild(NombreRazon);
            #endregion 

            XmlElement NIF = G.XmlDocument.CreateElement("sii", "NIF", G.SII); // NIF del emisor de la factura, empresa Rosell
            NIF.InnerText = _diccionarioValores[2];
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
            TipoNoExenta.InnerText = "S1"; //tipoNoExenta
            NoExenta.AppendChild(TipoNoExenta);

            XmlElement DesgloseIVA = G.XmlDocument.CreateElement("sii", "DesgloseIVA", G.SII);
            NoExenta.AppendChild(DesgloseIVA);

            #region DetalleIVA
            BucleDetalleIva(DesgloseIVA);
            #endregion

            XmlDocumentFragment frag = G.XmlDocument.CreateDocumentFragment();
            frag.AppendChild(DesgloseFactura);

            return frag;
        }

        private static void BucleDetalleIva(XmlElement DesgloseIVA)
        {
            for (int i = 4; i < 27; i += 5)
            {
                dynamic tipoImpositivo = _diccionarioValores[i +1];
                dynamic baseImponible = _diccionarioValores[i];
                dynamic cuotaRepercutida = _diccionarioValores[i + 2];

                if (tipoImpositivo != "")
                    DesgloseIVA.AppendChild(DetalleIva.XmlDetalleIva(tipoImpositivo, baseImponible, cuotaRepercutida));
            }
        }

        private static string ReemplazarAmpersandXML()
        {
            return string.Empty;
        }
    }
}
