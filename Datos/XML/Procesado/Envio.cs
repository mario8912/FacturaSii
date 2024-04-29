using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using G = Entidades.utils.Global;

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

        public void Request()
        {
            var resultado = "";

            Task tarea = new Task(() =>
            {
               resultado = LeerContenidoXML().Result;
            });
            tarea.Start();
            tarea.Wait();

            _contenido = new StringContent(resultado);

            ConfigurarHandler();

            using (HttpClient cliente = new HttpClient(_handler))
            {
                _respuestaServer = cliente.PostAsync(G.RutaEnvioPruebas, _contenido).Result;
                TryGuardarRespuesa();
            }
        }

        private async Task<string> LeerContenidoXML()
        {
            return await new StreamReader(G.RutaGuardarXmlEnvio, Encoding.UTF8).ReadToEndAsync();
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
            _reader?.Dispose();
            GC.SuppressFinalize(this);
            GC.Collect();
        }

        ~Envio()
        {
            Dispose();
        }
    }
}
