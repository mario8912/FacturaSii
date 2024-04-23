using Datos.XML.Procesado;

namespace Negocio.NegocioXML
{
    public class RespuestaXML
    {
        public void ProcesarRespuesta()
        {
            new Respuesta().RespuestaEstadoEnvio();
        }
    }
}
