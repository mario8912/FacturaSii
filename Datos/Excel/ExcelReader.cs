using Entidades.utils;
using System.Collections.Generic;


namespace Datos.Excel
{
    public class ExcelReader
    {
        private static Dictionary<int, dynamic> _diccionarioValores;
        private readonly Helper _listas;
        private readonly Excel _excel;

        public ExcelReader()
        {
            _listas = new Helper();
            _excel = new Excel();
        }

        public IEnumerable<Dictionary<int, dynamic>> LeerExcel(EventoProgreso eventoProgreso)
        {
            
            int rowCount = _excel.Rango.Rows.Count;
            eventoProgreso.ValorMaximoBarraProgreso = rowCount;

            for (int i = 2; i <= rowCount; i++)
            {
                _diccionarioValores = _listas.GetDiccionarioColumnasExcel();
                var tempDic = new Dictionary<int, dynamic>(_diccionarioValores);

                foreach (KeyValuePair<int, dynamic> item in _diccionarioValores)
                {
                    var rango = _excel.Rango.Cells[i, item.Key];

                    if (rango != null && rango.Value2 != null)
                        tempDic[item.Key] = rango.Value2.ToString();
                }

                eventoProgreso.AumentarProgreso();

                _diccionarioValores = tempDic;
                yield return _diccionarioValores;
            }
        }

        ~ExcelReader()
        {
            _excel.Dispose();
        }
    }
}
