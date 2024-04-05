using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Remoting.Messaging;
using System.Threading;
using System.Windows;
using System.Windows.Forms;

namespace Entidades.utils
{
    public class Helper
    {
        public static List<int> listaBaseCuotaTipo = new List<int> { 4, 6, 9, 11, 14, 16, 19, 21, 24, 26 };

        //la clave del diccionario hace referencia al indice de la columna del excel dond ese encuentran los datos deseados, 
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

                { 9, null }, //b2
                { 10, null }, //t2
                { 11, null }, //c2

                { 14, null }, //b3
                { 15, null }, //t3
                { 16, null }, //c3

                { 19, null }, //b4
                { 20, null }, //t4
                { 21, null }, //c4

                { 24, null }, //b5
                { 25, null }, //t5
                { 26, null }  //c5
            };
        }

        public static string SumaBases(Dictionary<int, dynamic> diccionarioValores)
        {
            float suma = 0;

            foreach (KeyValuePair<int, dynamic> itemDiccionario in diccionarioValores)
                if (listaBaseCuotaTipo.Contains(itemDiccionario.Key) && itemDiccionario.Value != null && itemDiccionario.Value != "")
                    suma += TryParseFloat("as");
                    //suma += TryParseFloat(itemDiccionario.Value);

            return suma.ToString();
        }

        private static float TryParseFloat(dynamic valor)
        {
            return GestorErrores.TryParseFloat(valor);
        }

        public static string FormatoEjercicio(string fecha)
        {
            return fecha.Substring(6, 4);
        }

        public static string FormatoPeriodo(string fecha)
        {
            return fecha.Substring(3, 2);

        }
    }
}
