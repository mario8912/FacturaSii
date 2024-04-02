using System;
using System.Xml;
using System.Xml.Schema;

namespace Datos.XML
{
    public class Validacion
    {
        private readonly string _rutaXsd = @"https://www2.agenciatributaria.gob.es/static_files/common/internet/dep/aplicaciones/es/aeat/ssii/fact/ws/SuministroInformacion.xsd";
        private XmlReaderSettings _settings;

        public Validacion()
        {
            CrearEsquemasYSettings();
            Validar();
        }

        private void CrearEsquemasYSettings()
        {
            XmlSchemaSet schemas = new XmlSchemaSet();
            schemas.Add(null, _rutaXsd);

            _settings = new XmlReaderSettings
            {
                ValidationType = ValidationType.Schema,
                Schemas = schemas
            };
        }

        private void Validar()
        {
            if(TryValidar())
                Console.WriteLine("El XML es válido según el XSD.");
            else
                Console.WriteLine("El XML no es válido según el XSD.");
        }

        private bool TryValidar()
        {
            try
            {
                using (XmlReader reader = XmlReader.Create(Entidades.utils.Global.RutaGuardarXml, _settings))
                    while (reader.Read()){}

                return true;
            }
            catch (XmlSchemaValidationException ex)
            {
                Console.WriteLine($"Error de validación XML: {ex.Message}");
                return false;
            }
            catch (XmlException ex)
            {
                Console.WriteLine($"Error de XML: {ex.Message}");
                return false;
            }
        }
    }
}
