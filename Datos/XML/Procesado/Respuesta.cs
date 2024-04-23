using System;
using O = System.Data.OleDb;
using System.Data;
using System.Diagnostics;
using System.Xml;
using G = Entidades.utils.Global;

public class Respuesta
{
    private readonly XmlDocument _xmlDoc;
    private XmlNamespaceManager _nsMgr;
    private string _xmlPath;
    private DataTable _tabla;
    public Respuesta()
    {
        _xmlDoc = new XmlDocument();
        _xmlDoc.Load(G.RutaGuardarXmlRespuesta);

        ConfiguracionNamespace();
        CrearTabla();
    }

    private XmlNode ConfiguracionNamespace()
    {
        _xmlPath = "env:Envelope/env:Body/siiR:RespuestaLRFacturasEmitidas/siiR:EstadoEnvio";

        _nsMgr = new XmlNamespaceManager(_xmlDoc.NameTable);
        _nsMgr.AddNamespace("env", G.SOAPENV);
        _nsMgr.AddNamespace("sii", G.SII);
        _nsMgr.AddNamespace("siiR", G.SII_R);

        return _xmlDoc.SelectSingleNode(_xmlPath, _nsMgr);
    }

    private string RespuestaEstadoEnvio()
    {
        XmlNode nodo = ConfiguracionNamespace();
        return nodo.InnerText; //try
    }

    private void CrearTabla()
    {
        _tabla = new DataTable();
        DataColumn colIdFactura = new DataColumn("IdFactura");
        DataColumn colEstadoReg = new DataColumn("EstadoRegistro");
        DataColumn colCodError = new DataColumn("CodigoError");
        DataColumn colMensajeError = new DataColumn("MensajeError");

        _tabla.Columns.Add(colIdFactura);
        _tabla.Columns.Add(colEstadoReg);
        _tabla.Columns.Add(colCodError);
        _tabla.Columns.Add(colMensajeError);
    }

    public DataTable Tabla()
    {
        Stopwatch sp = new Stopwatch();
        sp.Start();

        _xmlPath = "env:Envelope/env:Body/siiR:RespuestaLRFacturasEmitidas/siiR:RespuestaLinea";
        XmlNodeList node = _xmlDoc.SelectNodes(_xmlPath, _nsMgr);

        foreach (XmlNode nodoFacturas in node)
        {
            string idFactura = nodoFacturas.ChildNodes[0].ChildNodes[1].InnerText;
            string estadoRegistro = nodoFacturas.ChildNodes[1].InnerText;
            string codigoError = string.Empty;
            string descripcionError = string.Empty;
            string csv = string.Empty;
            string estadoDuplicado = string.Empty;

            if (estadoRegistro != "Correcto")
            {
                codigoError = nodoFacturas.ChildNodes[2].InnerText;
                descripcionError = nodoFacturas.ChildNodes[3].InnerText;
                if (descripcionError == "Factura duplicada")
                {
                    csv = nodoFacturas.ChildNodes[4].InnerText;
                    estadoDuplicado = nodoFacturas.ChildNodes[5].ChildNodes[0].InnerText;
                }
            }


            DataRow fila = _tabla.NewRow();
            fila["IdFactura"] = idFactura;
            fila["EstadoRegistro"] = estadoRegistro;
            fila["CodigoError"] = codigoError;
            fila["MensajeError"] = descripcionError;
            _tabla.Rows.Add(fila);
        }
        sp.Stop();
        Console.WriteLine(sp.Elapsed.ToString());

        _tabla.TableName = "Facturas";
        _tabla.WriteXml("datos.xml");

        return _tabla;
    }
}