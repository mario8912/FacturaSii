﻿using G = Entidades.utils.Global;
using Microsoft.Office.Interop.Excel;
using System;
using System.Runtime.InteropServices;

namespace Datos.Excel
{
    internal class Excel : IDisposable
    {
        public Application ExcelApp { get; private set; }
        public Workbook Libro { get; private set; }
        public Worksheet Hoja { get; private set; }
        public Range Rango { get; private set; }

        public Excel()
        {
            ExcelApp = new Application();
            Libro = ExcelApp.Workbooks.Open(G.ExcelFile);
            Hoja = Libro.Sheets[1];
            Rango = Hoja.UsedRange;
        }

        public void Dispose()
        {
            Libro?.Close();  
            ExcelApp?.Quit();
            TryLimpiarRecursos();
        }

        private void TryLimpiarRecursos()
        {
            try
            {
                LimpiarRecursos();
            }
            catch (Exception)
            {

                throw new Exception();
            }
            
        }

        private void LimpiarRecursos()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(Rango);
            Marshal.ReleaseComObject(Hoja);
            Marshal.ReleaseComObject(Libro);
            Marshal.ReleaseComObject(ExcelApp);

        }
    }
}
