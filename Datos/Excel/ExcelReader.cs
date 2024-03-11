using Entidades.utils;
using Entidades.utils.XML;  
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using uExcel = Microsoft.Office.Interop.Excel;

namespace Datos.Excel
{
    public class ExcelReader
    {
        private static Dictionary<int, TipoValor> _diccionarioValores;

        private static uExcel.Application _excelApp;
        private static Workbook _libro;
        private static Worksheet _hoja;
        private static Range _rango;

        public static void LeerExcel(string filePath)
        {
            Listas listas = new Listas();

            _excelApp = new uExcel.Application();
            _libro = _excelApp.Workbooks.Open(filePath);
            _hoja = _libro.Sheets[1];
            _rango = _hoja.UsedRange;

            int rowCount = _rango.Rows.Count;



            for (int i = 2; i <= 2; i++)
            {
                _diccionarioValores = listas.DiccionarioCeldas();

                foreach (var item in _diccionarioValores)
                {
                    var rango = _rango.Cells[i, item.Key];

                    if (rango != null && rango.Value2 != null)
                    {
                        item.Value.Valor = rango.Value2.ToString();

                        NegocioXml().CrearXml(_diccionarioValores);
                    }
                }   
            }

            LimpiarRecursos();
        }
        
        private static void LimpiarRecursos()
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
