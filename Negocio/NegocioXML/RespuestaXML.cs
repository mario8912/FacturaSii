using Datos.XML.Procesado;
using System.Collections;
using System.Data;

namespace Negocio.NegocioXML
{
    public class RespuestaXML
    {
        private string _respuesta;
        public DataTable ProcesarRespuesta()
        {
            return new Respuesta().Tabla();
        }

        private bool SwitchRespuesta()
        {
            bool respuesta = false;

            switch (_respuesta)
            {
                case "Incorrecto":
                    break;
                case "Correcto":
                    respuesta = true;
                    break;
                case "ParcialmenteCorrecto":
                    break;

                default:
                    break;
            }

            return respuesta;
        }
    }
}
