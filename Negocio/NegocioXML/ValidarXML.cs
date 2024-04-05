using Datos.XML;
using System.Security.AccessControl;

namespace Negocio.NegocioXML
{
    public class ValidarXML
    {
        private static readonly IValidacion validacion = new Validacion();
        public static string ValidarXml()
        {
            return validacion.Validar();
        }
    }
}


