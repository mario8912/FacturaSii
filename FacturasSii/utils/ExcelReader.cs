using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace FacturasSii.Utils
{
    public class ExcelReader
    {
        public static void ReadExcel(string filePath)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    // New line
                    if (j == 1)
                        Console.Write("\r\n");

                    // Write value to the console
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");
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
        }
    }
}
