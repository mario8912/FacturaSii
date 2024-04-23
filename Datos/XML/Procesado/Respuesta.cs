using System.Xml;
using G = Entidades.utils.Global;

public class Respuesta
{
    private readonly XmlDocument _xmlDoc;
    public Respuesta()
    {
        _xmlDoc = new XmlDocument();
        _xmlDoc.Load(G.RutaGuardarXmlRespuesta);
    }

    public string RespuestaEstadoEnvio()
    {
        string xmlPath = "env:Envelope/env:Body/siiR:RespuestaLRFacturasEmitidas/siiR:EstadoEnvio";
        
        XmlNamespaceManager nsManager = new XmlNamespaceManager(_xmlDoc.NameTable);
        nsManager.AddNamespace("env", G.SOAPENV);
        nsManager.AddNamespace("sii", G.SII);
        nsManager.AddNamespace("siiR", G.SII_R);

        XmlNode node = _xmlDoc.SelectSingleNode(xmlPath, nsManager);

        return node.InnerText;
    }
}