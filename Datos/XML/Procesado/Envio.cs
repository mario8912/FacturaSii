using System;
using System.Net.Http;
using System.Text;
using System.IO;
using System.Diagnostics;

namespace Datos.XML.Procesado
{
    public class Envio : /*IEnvio,*/ IDisposable
    {
        private readonly HttpClient _cliente;
        private StringContent _contenido;
        private HttpResponseMessage _respuestaServer;
        private string _contenidoRespuesta;

        public Envio()
        {
            _cliente = new HttpClient();
        }

        public async void Request(string url, string content)
        {
            _contenido = new StringContent(content, Encoding.UTF8, "application/xml");
            _respuestaServer = _cliente.PostAsync(url, _contenido).Result;

            if (_respuestaServer.IsSuccessStatusCode)
            {
                //Correcto
                _contenidoRespuesta = await _respuestaServer.Content.ReadAsStringAsync();
                File.WriteAllText("respuesta.html", _contenidoRespuesta);
                Process.Start("respuesta.html");
            }
            else
            {
                //Incorrecto
                Console.WriteLine("Error al enviar XML: " + _respuestaServer);
            }

        }

        public void Dispose()
        {
            GC.SuppressFinalize(this);
            GC.Collect();
            _cliente.Dispose();
        }

        ~Envio()
        {
            Dispose();
        }
    }
}
