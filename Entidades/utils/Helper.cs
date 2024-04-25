using System;
using System.Collections.Generic;
using System.IO;

namespace Entidades.utils
{
    public class Helper
    {
        public static List<int> listaBaseCuotaTipo = new List<int> { 4, 6, 8, 9, 11, 13, 14, 16, 18, 19, 21, 23, 24, 26, 28 };

        //la clave del diccionario hace referencia al indice de la columna del excel donde ese encuentran los datos deseados, 
        //ya que siempre están en la misma posición pero de forma desordenada, hay columnas vacías, etc.
        //el valor respectivo contiene el propio valor de la casilla a la que la clave hace refrencia, forma así un diccionaro con el que podemos
        //iterar sobre las filas del excel usando la clave como índicie y guardando el valor de la casilla en el valor del diccionario .
        public Dictionary<int, dynamic> GetDiccionarioColumnasExcel()
        {
            return new Dictionary<int, dynamic>
            {
                { 0, null }, //num factura
                { 1, null }, //fecha expedicion
                { 2, null }, //nifid
                { 3, null }, //nombre razon

                { 4, null }, //b1
                { 5, null }, //t1
                { 6, null }, //c1
                { 7, null }, //tre1
                { 8, null }, //cre1
               
                { 9, null }, //b2
                { 10, null }, //t2
                { 11, null }, //c2
                { 12, null }, //tre2
                { 13, null }, //cre2

                { 14, null }, //b3
                { 15, null }, //t3
                { 16, null }, //c3
                { 17, null }, //tre3
                { 18, null }, //cre3

                { 19, null }, //b4
                { 20, null }, //t4
                { 21, null }, //c4
                { 22, null }, //tre4
                { 23, null }, //cre4

                { 24, null }, //b5
                { 25, null }, //t5
                { 26, null }, //c5
                { 27, null }, //tre5
                { 28, null }  //cre5
            };
        }

        public static string SumaBases(Dictionary<int, dynamic> diccionarioValores)
        {
            float suma = 0f;

            foreach (KeyValuePair<int, dynamic> itemDiccionario in diccionarioValores)
            {
                if (listaBaseCuotaTipo.Contains(itemDiccionario.Key) && itemDiccionario.Value != null && itemDiccionario.Value != "")
                {
                    suma += TryParseFloat(itemDiccionario.Value);
                }
            }
                
            return ReemplazarComaPunto(suma.ToString("0.00"));
        }

        public static string ReemplazarComaPunto(dynamic val)
        {
            return val.ToString().Replace(',', '.');
        }

        private static float TryParseFloat(dynamic valor)
        {
            return GestorErrores.TryParseFloat(valor);
        }

        public static void SetHora()
        {
            Global.FechaGuardado = DateTime.Now.ToString("yy_MM_dd_HH_mm_ss_ffff");
        }
    }
}
