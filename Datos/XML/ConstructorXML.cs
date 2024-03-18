using G = Entidades.utils.Global;
using Entidades.utils.XML;
using System.Xml;
using System.Collections.Generic;
using System;
using System.IO;
using System.Linq;

namespace Datos.XML
{
    public class ConstructorXML : IConstructorXML
    {
        private XmlElement _suministroLR;
        private readonly Envoltorio _envoltorio = new Envoltorio();

        public ConstructorXML EstructuraXML()
        {
            _envoltorio.EstructuraPrincipalXML();
            _suministroLR = Envoltorio.SuministroLR;
            return this;
        }

        public ConstructorXML EstructuraCabeceraXML()
        {
            _suministroLR.AppendChild(Cabecera.CabeceraXml());
            return this;
        }

        public void EstructuraFacturaXML(IEnumerable<Dictionary<int, dynamic>> diccionarioValores)
        {
            diccionarioValores.ToList().ForEach(item => _suministroLR.AppendChild(FacturaEmitida.XmlFactura(item)));
        }

        public void GuardarXML()
        {
            try
            {
                BorrarXmlAntiguo();
                G.XmlDocument.Save(G.RutaGuardarXml);
            }
            catch (Exception)
            {
                throw new Exception("Error al guardar el archivo XML");
            }
        }

        private void BorrarXmlAntiguo()
        {
            if (Directory.Exists(G.RutaGuardarXml))
                File.Delete(G.RutaGuardarXml);
        }
    }
}
