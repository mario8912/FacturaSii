using Entidades.utils;
using System.Collections.Generic;


namespace Datos.Excel
{
    public class ExcelReader
    {
        private static Dictionary<int, dynamic> _diccionarioValores;
        private Listas _listas;

        public IEnumerable<Dictionary<int, dynamic>> LeerExcel()
        {
            _listas = new Listas(); 

            using (Excel excel = new Excel())
            {
                int rowCount = excel.Rango.Rows.Count;

                for (int i = 2; i <= 20; i++)
                {
                    _diccionarioValores = _listas.GetDiccionarioColumnasExcel();
                    var tempDic = new Dictionary<int, dynamic>(_diccionarioValores);

                    foreach (KeyValuePair<int, dynamic> item in _diccionarioValores)
                    {
                        var rango = excel.Rango.Cells[i, item.Key];

                        if (rango != null && rango.Value2 != null)
                            tempDic[item.Key] = rango.Value2.ToString();
                    }

                    _diccionarioValores = tempDic;
                    yield return _diccionarioValores;
                }
            }
        }
    }
}
