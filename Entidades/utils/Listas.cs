using System.Collections.Generic;

namespace Entidades.utils
{
    public class Listas
    {
        public static string[] stringsAComparar = {"Base1", "Cuota1", "Base2", "Cuota2", "Base3", "Cuota3", "Base4", "Cuota4", "Base5", "Cuota5"};

        public Dictionary<int, TipoValor> DiccionarioCeldas()
        {
            return new Dictionary<int, TipoValor>
            {
                { 1, new TipoValor("NumFactura") },
                { 2, new TipoValor("FechaExpedicion") },
                { 3, new TipoValor("NIFID") },
                { 4, new TipoValor("NombreRazon") },

                { 5, new TipoValor("Base1") },
                { 6, new TipoValor("Tipo1") },
                { 7, new TipoValor("Cuota1") },

                { 10, new TipoValor("Base2") },
                { 11, new TipoValor("Tipo2") },
                { 12, new TipoValor("Cuota2") },

                { 15, new TipoValor("Base3") },
                { 16, new TipoValor("Tipo3") },
                { 17, new TipoValor("Cuota3") },

                { 20, new TipoValor("Base4") },
                { 21, new TipoValor("Tipo4") },
                { 22, new TipoValor("Cuota4") },

                { 25, new TipoValor("Base5") },
                { 26, new TipoValor("Tipo5") },
                { 27, new TipoValor("Cuota5") }
            };
            
        }
    }

    public class TipoValor
    {
        public string Campo { get; private set; }
        public string Valor { get; set; }
        
        public TipoValor(string campo, string valor = null)
        {
            Campo = campo;
            Valor = valor;
        }
    }

}
