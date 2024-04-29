using System;
using System.Net.Http;
using System.IO;
using System.Diagnostics;
using System.Security.Cryptography.X509Certificates;
using G = Entidades.utils.Global;
using System.Text;
using System.Threading.Tasks;

namespace Datos.XML.Procesado
{
    public class Envio : IDisposable
    {
        private readonly HttpClientHandler _handler;
        private X509Certificate2 _certificate;
        private StringContent _contenido;
        private HttpResponseMessage _respuestaServer;
        private string _contenidoRespuesta;
        private StreamReader _reader;
        
        public Envio()
        {
            _handler = new HttpClientHandler();
        }

        public async void Request()
        {
            _contenido  = new StringContent(await LeerContenidoXML().Start());

            ConfigurarHandler();
            
            using (HttpClient cliente = new HttpClient(_handler))
            {
                _respuestaServer = cliente.PostAsync(G.RutaEnvioPruebas, _contenido).Result;

                TryGuardarRespuesa();
            }
        }

        private Task LeerContenidoXML()
        {
            _reader = new StreamReader(G.RutaGuardarXmlEnvio, Encoding.UTF8);
            return new Task(() => {
                _reader.ReadToEndAsync();
            });
        }

        private async void TryGuardarRespuesa()
        {
            var nombreRespuesta = GuardarRespuesta();
            if (_respuestaServer.IsSuccessStatusCode)
            {
                _contenidoRespuesta = await _respuestaServer.Content.ReadAsStringAsync();
                File.WriteAllText(nombreRespuesta, _contenidoRespuesta);
                Process.Start(nombreRespuesta);
            }
            else
                Console.WriteLine("Error al enviar XML: " + _respuestaServer);
        }

        

        private void ConfigurarHandler()
        {
            var rutaCertificado = Path.Combine(G.RutaAppExe, @"DISROSELLSL.pfx");

            _certificate = new X509Certificate2(rutaCertificado, "1234");
            _handler.ClientCertificates.Add(_certificate);
        }

        private string GuardarRespuesta()
        {
            var nombreXmlRespuesta = string.Format("R-{0}.xml", G.FechaGuardado);
            return G.RutaGuardarXmlRespuesta = Path.Combine(G.RutaDirectorioData, nombreXmlRespuesta);
        }

        public void Dispose()
        {
            _reader.Dispose();
            GC.SuppressFinalize(this);
            GC.Collect();
        }

        ~Envio()
        {
            Dispose();
        }
    }
}
