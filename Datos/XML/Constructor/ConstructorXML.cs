using G = Entidades.utils.Global;
using Entidades.utils.XML;
using System.Xml;
using System.Collections.Generic;
using System;
using System.IO;
using Entidades.utils;

namespace Datos.XML
{
    public class ConstructorXML
    {
        private XmlElement _suministroLR;
        private readonly Envoltorio _envoltorio = new Envoltorio();
        private IEnumerable<Dictionary<int, dynamic>> _diccionarioValores;

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

        public bool TryEstructuraFacturaXML(IEnumerable<Dictionary<int, dynamic>> diccionarioValores)
        {
            _diccionarioValores = diccionarioValores;
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

        public void TryGuardarXML()
        {
            try
            {
                BorrarXmlAntiguo();
                GuardarXML();
                
            }
            catch (Exception ex )
            {
                throw new Exception("Error al guardar el archivo XML" +
                    $"{Environment.NewLine}{ex.Message}");
            }
        }

        private void BorrarXmlAntiguo()
        {
            if (Directory.Exists(G.RutaGuardarXmlEnvio))
                File.Delete(G.RutaGuardarXmlEnvio);
        }

        private void GuardarXML()
        {
            GetHora();
            GuardarXmlEnvio();
            
        }

        private void GuardarXmlEnvio()
        {
            string rutaGuardado = Helper.GetRutaGuardadoXml("envio");
            Console.WriteLine(rutaGuardado);
            G.XmlDocument.Save(rutaGuardado);
        }

        private void GetHora()
        {
            G.FechaGuardado = DateTime.Now.ToString("yy_MM_dd_HH_mm_ss_ffff");
        }
    }
}
