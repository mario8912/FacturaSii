﻿using System;
using O = System.Data.OleDb;
using System.Data;
using System.Diagnostics;
using System.Xml;
using G = Entidades.utils.Global;

public class Respuesta
{
    private readonly XmlDocument _xmlDoc;
    private XmlNamespaceManager _namespaceManager;
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

        _namespaceManager = new XmlNamespaceManager(_xmlDoc.NameTable);
        _namespaceManager.AddNamespace("env", G.SOAPENV);
        _namespaceManager.AddNamespace("sii", G.SII);
        _namespaceManager.AddNamespace("siiR", G.SII_R);

        return _xmlDoc.SelectSingleNode(_xmlPath, _namespaceManager);
    }

    private void CrearTabla()
    {
        _tabla = new DataTable();
        DataColumn colIdFactura = new DataColumn("IdFactura");
        DataColumn colEstadoReg = new DataColumn("EstadoRegistro");
        DataColumn colCodError = new DataColumn("CodigoError");
        DataColumn colMensajeError = new DataColumn("MensajeError");
        DataColumn colCsv = new DataColumn("CSV");
        DataColumn codEstadoDuplicado = new DataColumn("EstadoDuplicado");

        _tabla.Columns.Add(colIdFactura);
        _tabla.Columns.Add(colEstadoReg);
        _tabla.Columns.Add(colCodError);
        _tabla.Columns.Add(colMensajeError);
        _tabla.Columns.Add(colCsv);
        _tabla.Columns.Add(codEstadoDuplicado);
    }

    public DataTable Tabla()
    {
        _xmlPath = "env:Envelope/env:Body/siiR:RespuestaLRFacturasEmitidas/siiR:RespuestaLinea";
        XmlNodeList node = _xmlDoc.SelectNodes(_xmlPath, _namespaceManager);

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

                if (codigoError == "3000")
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
            fila["CSV"] = csv;
            fila["EstadoDuplicado"] = estadoDuplicado;

            _tabla.Rows.Add(fila);
        }

        _tabla.TableName = "Facturas";
        _tabla.WriteXml("datos.xml");

        return _tabla;
    }
}