using G = Entidades.utils.Global;
using Entidades.utils.XML;
using System.Xml;
using System.Collections.Generic;
using System;
using System.IO;

namespace Datos.XML
{
    public class ConstructorXML
    {
        private XmlElement _suministroLR;
        private readonly Envoltorio _envoltorio = new Envoltorio();
        private readonly IEnumerable<Dictionary<int, dynamic>> _diccionarioValores;

        public ConstructorXML(IEnumerable<Dictionary<int, dynamic>> diccionarioValores) 
        {
            _diccionarioValores = new List<Dictionary<int, dynamic>>(diccionarioValores);
        }

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

        public bool TryEstructuraFacturaXML()
        {
            try
            {
                BucleEstructuraFacturaXML();
                return true;
            }
            catch (Exception ex)
            {
                new Exception($"Error al crear la estructura XML del Registro de faccturas.{Environment.NewLine} {ex.Message} ");
                return false;
            }
        }

        private void BucleEstructuraFacturaXML()
        {
            foreach (var item in _diccionarioValores)
                _suministroLR.AppendChild(FacturaEmitida.XmlFactura(item));
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
