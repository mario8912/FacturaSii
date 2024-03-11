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
        private Dictionary<int, TipoValor> _diccionarioValores;

        public void LeerExcel(string filePath)
        {
            Listas listas = new Listas();
            uExcel.Application xlApp = new uExcel.Application();
            Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
            _Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;

            for (int i = 2; i <= 2; i++)
            {
                _diccionarioValores = listas.DiccionarioCeldas();

                foreach (var item in _diccionarioValores)
                {
                    if (xlRange.Cells[i, item.Key] != null && xlRange.Cells[i, item.Key].Value2 != null)
                        item.Value.Valor = xlRange.Cells[i, item.Key].Value2.ToString();
                }   
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);


            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);
            
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            
            GC.Collect();

            CrearXml();
        }
        
        public void CrearXml()
        {

            Envoltorio.EstructuraExternaXml();
            MessageBox.Show("Archivo creado");
        }
    }
}
