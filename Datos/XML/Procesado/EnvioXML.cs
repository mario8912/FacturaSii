using System;
using System.Net.Http;
using System.IO;
using System.Diagnostics;
using System.Security.Cryptography.X509Certificates;
using G = Entidades.utils.Global;
using H = Entidades.utils.Helper;
using System.Text;
using System.Threading.Tasks;

namespace Datos.XML.Procesado
{
    public class EnvioXML : IDisposable
    {
        private readonly HttpClientHandler _handler;
        private X509Certificate2 _certificate;

        private StringContent _httpContenido;
        private HttpResponseMessage _respuestaServer;
        private string _contenidoRespuesta;
        

        public EnvioXML()
        {
            _handler = new HttpClientHandler();
        }

        public async void Request()
        {
            string respuesta = await Lectura();
            _httpContenido = new StringContent(respuesta);

            ConfigurarHandler();

            await aaaa();



        }

        private async Task aaaa()
        {
            using (HttpClient cliente = new HttpClient(_handler))
            {
                _respuestaServer = cliente.PostAsync(G.RutaEnvioPruebas, _httpContenido).Result;

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
        }

        private async Task<string> Lectura()
        {
            using (StreamReader reader = new StreamReader(G.RutaGuardarXmlEnvio, Encoding.UTF8))
            {
                return await reader.ReadToEndAsync();
            }
        }

        private string GuardarRespuesta()
        {
            var nombreXmlRespuesta = string.Format("R-{0}.xml", G.FechaGuardado);
            return G.RutaGuardarXmlRespuesta = Path.Combine(G.RutaDirectorioData, nombreXmlRespuesta);
        }

        private void ConfigurarHandler()
        {
            var rutaCertificado = Path.Combine(G.RutaAppExe, @"DISROSELLSL.pfx");
            
            _certificate = new X509Certificate2(rutaCertificado, "1234");
            _handler.ClientCertificates.Add(_certificate);
        }

        public void Dispose()
        {
            _handler?.Dispose();
            GC.SuppressFinalize(this);
            GC.Collect();
        }

        ~EnvioXML()
        {
            Dispose();
        }
    }
}
