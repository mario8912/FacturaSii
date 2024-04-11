﻿using System;
using System.Net.Http;
using System.IO;
using System.Diagnostics;
using System.Security.Cryptography.X509Certificates;
using G = Entidades.utils.Global;
using System.Text;

namespace Datos.XML.Procesado
{
    public class Envio : IDisposable
    {
        private readonly HttpClientHandler _handler;
        private X509Certificate2 _certificate;

        private StringContent _contenido;
        private HttpResponseMessage _respuestaServer;
        private string _contenidoRespuesta;

        public Envio()
        {
            _handler = new HttpClientHandler();
        }

        public async void Request()
        {
            var xmlFilePath = G.RutaGuardarXml;
            //var xmlFilePath = @"E:\mipc\escritorio\FacturaSii\Entidades\utils\XML\factura.xml";

            string xmlContent;
            using (StreamReader reader = new StreamReader(xmlFilePath, Encoding.UTF8))
            {
                xmlContent =  reader.ReadToEnd();
            }
            _contenido  = new StringContent(xmlContent);

            ConfigurarHandler();

            using (HttpClient cliente = new HttpClient(_handler))
            {
                _respuestaServer = cliente.PostAsync(G.RutaEnvioPruebas, _contenido).Result;

                if (_respuestaServer.IsSuccessStatusCode)
                {
                    //Correcto
                    _contenidoRespuesta = await _respuestaServer.Content.ReadAsStringAsync();
                    File.WriteAllText("respuesta.xml", _contenidoRespuesta);
                    Process.Start("respuesta.xml");
                }
                else
                    Console.WriteLine("Error al enviar XML: " + _respuestaServer);

            }
        }

        private void ConfigurarHandler()
        {
            var rutaCertificado = Path.Combine(G.RutaAppExe, @"DISROSELLSL.pfx");
            
            _certificate = new X509Certificate2(rutaCertificado, "1234");
            _handler.ClientCertificates.Add(_certificate);
        }

        public void Dispose()
        {
            GC.SuppressFinalize(this);
            GC.Collect();
        }

        ~Envio()
        {
            Dispose();
        }
    }
}
