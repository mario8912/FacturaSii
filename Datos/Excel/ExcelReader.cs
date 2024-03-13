using Entidades.utils;
using System.Collections.Generic;
using ProgressBar = Entidades.utils;


namespace Datos.Excel
{
    public class ExcelReader
    {
        private static Dictionary<int, dynamic> _diccionarioValores;
        private Helper _listas;
        private EventoProgreso _eventoProgreso;

        public IEnumerable<Dictionary<int, dynamic>> LeerExcel()
        {
            _eventoProgreso = new EventoProgreso();
            _listas = new Helper(); 

            using (Excel excel = new Excel())
            {
                int rowCount = excel.Rango.Rows.Count;
                _eventoProgreso.ValorMaximoBarraProgreso = rowCount;

                for (int i = 2; i <= rowCount; i++)
                {
                    _diccionarioValores = _listas.GetDiccionarioColumnasExcel();
                    var tempDic = new Dictionary<int, dynamic>(_diccionarioValores);

                    foreach (KeyValuePair<int, dynamic> item in _diccionarioValores)
                    {
                        var rango = excel.Rango.Cells[i, item.Key];

                        if (rango != null && rango.Value2 != null)
                            tempDic[item.Key] = rango.Value2.ToString();
                    }

                    _eventoProgreso.AumentarProgreso();

                    _diccionarioValores = tempDic;
                    yield return _diccionarioValores;
                }
            }
        }
    }
}
