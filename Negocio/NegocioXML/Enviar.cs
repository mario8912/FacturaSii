using Datos.XML.Procesado;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using G = Entidades.utils.Global;

namespace Negocio.NegocioXML
{
    public class Enviar
    {
        public void Envio()
        {
            new EnvioXML().Request();
        }
    }
}
