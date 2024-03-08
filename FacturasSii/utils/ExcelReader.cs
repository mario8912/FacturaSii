using FacturasSii.utils;
using Microsoft.Office.Interop.Excel;
using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace FacturasSii.Utils
{
    public class ExcelReader
    {
        public static void ReadExcel(string filePath)
        {
            Application xlApp = new Excel.Application();
            Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
            _Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            //int columnas = 7;

            for (int i = 2; i <= rowCount; i++)
            {
                foreach(var item in Listas.diccionarioCeldas)
                {
                    if (xlRange.Cells[i, item.Key] != null && xlRange.Cells[i, item.Key].Value2 != null)
                    {
                        item.Value.Valor = xlRange.Cells[i, item.Key].Value2.ToString();
                    }
                }
            }



            // Cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();
            // Close and release
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            // Close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);
            // Quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            // Cleanup
            GC.Collect();

            //generate a function that inserts the values of the class inside the dictionary into the xml  file


        }
    }
}
