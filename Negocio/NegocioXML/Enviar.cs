using Datos.XML.Procesado;
using G = Entidades.utils.Global;

namespace Negocio.NegocioXML
{
    public class Enviar
    {
        public static void Envio()
        {
            Envio en = new Envio();
            en.Request(G.RutaEnvioPruebas ,G.RutaGuardarXml);
        }
    }
}
