using System;
using System.Collections.Generic;
using System.Xml.Schema;

namespace Entidades.utils
{
    public class Helper
    {
        public static List<int> listaBaseCuotaTipo = new List<int> { 4, 6, 9, 11, 14, 16, 19, 21, 24, 26 };
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

        public List<Dictionary<int, dynamic>> DiseccionarIvaEnListas(Dictionary<int, dynamic> diccionario)
        {
            List<Dictionary<int, dynamic>> listaIvas = new List<Dictionary<int, dynamic>>();
            Dictionary<int, dynamic> diccionarioIva;

            for (int i = 4; i <= 26; i += 5)
            {
                var value = diccionario[i];

                if (value[0] != null && value[0] != 0)
                {
                    diccionarioIva = new Dictionary<int, dynamic>
                    {
                        { i, null}, //base
                        { i + 1, null }, //tipo
                        { i + 2, null } //cuota
                    };

                    listaIvas.Add(diccionarioIva);
                }
            }

            return listaIvas;
        }

        public static string SumaBases(Dictionary<int, dynamic> diccionarioValores)
        {
            float suma = 0;

            foreach (KeyValuePair<int, dynamic> itemDiccionario in diccionarioValores)
                if (listaBaseCuotaTipo.Contains(itemDiccionario.Key) && itemDiccionario.Value != null && itemDiccionario.Value != "")
                {
                    suma += float.Parse(itemDiccionario.Value);
                }
            return suma.ToString();
        }
    }

}
