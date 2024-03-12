using Entidades.utils;
using G = Entidades.utils.Global;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Datos.Excel
{
    public class ExcelReader
    {

        private Dictionary<int, TipoValor> _diccionarioValores;
        private readonly IEnumerable<Dictionary<int, TipoValor>> _listaDiccionarios;

        private Application _excelApp;
        private Workbook _libro;
        private Worksheet _hoja;
        private Range _rango;

        public  IEnumerable<Dictionary<int, TipoValor>> LeerExcel()
        {
            Listas listas = new Listas();
            
            _excelApp = new Application();
            _libro = _excelApp.Workbooks.Open(G.ExcelFile);
            _hoja = _libro.Sheets[1];
            _rango = _hoja.UsedRange;

            int rowCount = _rango.Rows.Count;

            for (int i = 2; i <= rowCount; i++)
            {
                _diccionarioValores = listas.DiccionarioCeldas();

                foreach (var item in _diccionarioValores)
                {
                    var rango = _rango.Cells[i, item.Key];

                    if (rango != null && rango.Value2 != null)
                    {
                        item.Value.Valor = rango.Value2.ToString();
                    }
                }
                yield return _diccionarioValores;
            }
            LimpiarRecursos();
        }

        private void LimpiarRecursos()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(_rango);
            Marshal.ReleaseComObject(_hoja);


            _libro.Close();
            Marshal.ReleaseComObject(_libro);

            _excelApp.Quit();
            Marshal.ReleaseComObject(_excelApp);

            GC.Collect();
        }
    }
}
