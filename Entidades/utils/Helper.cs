using System.Collections.Generic;

namespace Entidades.utils
{
    public class Helper
    {
        public static List<int> listaBaseCuotaTipo = new List<int> {5, 6, 7, 10, 11, 12, 15, 16 ,17, 20, 21, 22, 25, 26, 27 };
        public Dictionary<int, dynamic> GetDiccionarioColumnasExcel()
        {
            return new Dictionary<int, dynamic>
            {
                { 1, null }, //num factura
                { 2, null }, //fecha expedicion
                { 3, null }, //nifid
                { 4, null }, //nombre razon

                { 5, null }, //b1
                { 6, null }, //t1
                { 7, null }, //c1

                { 10, null }, //b2
                { 11, null }, //t2
                { 12, null }, //c2

                { 15, null }, //b3
                { 16, null }, //t3
                { 17, null }, //c3

                { 20, null }, //b4
                { 21, null }, //t4
                { 22, null }, //c4

                { 25, null }, //b5
                { 26, null }, //t5
                { 27, null }  //c5
            };
            
        }

        public static string SumaBases(Dictionary<int, dynamic> diccionarioValores)
        {
            float suma = 0;

            foreach (KeyValuePair<int, dynamic> itemDiccionario in diccionarioValores)
                if (listaBaseCuotaTipo.Contains(itemDiccionario.Key) && itemDiccionario.Value != null)
                    suma += float.Parse(itemDiccionario.Value);

            return suma.ToString();
        }
    }

}
