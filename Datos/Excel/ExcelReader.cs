using G = Entidades.utils.Global;   
using Entidades.utils;
using OfficeOpenXml;
using System.Collections.Generic;
using DT = System.Data;
using System.Diagnostics;
using System;
using System.Data.OleDb;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Data;


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

            _excel.Dispose();
        }
    }

    public class ExcelReader1
    {
        private readonly Excel _excel;  
        private DT.DataTable _dataTable;

        public ExcelReader1()
        { 
            _dataTable = new DT.DataTable();
        }

        public void LeerExcel()
        {
            Stopwatch st = new Stopwatch();
            st.Start();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage(new System.IO.FileInfo(G.ExcelFile)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Selecciona la primera hoja del libro

                // Agregar las columnas al DataTable
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    _dataTable.Columns.Add(worksheet.Cells[1, col].Value.ToString());
                }

                // Agregar las filas al DataTable
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    DataRow newRow = _dataTable.Rows.Add();
                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        newRow[col - 1] = worksheet.Cells[row, col].Value;
                    }
                }
                
                _excel.Dispose();
            }
            st.Stop();

            Console.WriteLine("bucle:" + st.Elapsed);

            
        }

        public /*DT.DataTable*/ DataSet LeerExcelRs()
        {
            var sConnectionString = "" + 
            "Provider=Microsoft.Jet.OLEDB.4.0;" +
            "Data Source=" + G.ExcelFile + ";" +
            "Extended Properties=Excel 8.0;";

            using (new Excel())
            {
                using (OleDbConnection objConn = new OleDbConnection(sConnectionString))
                {
                    objConn.Open();

                    OleDbCommand objCmdSelect = new OleDbCommand("SELECT * FROM [Hoja1$]", objConn);

                    OleDbDataAdapter objAdapter1 = new OleDbDataAdapter
                    {
                        SelectCommand = objCmdSelect
                    };

                    DataSet objDataset1 = new DataSet();
                    objAdapter1.Fill(objDataset1);

                    objConn.Close();

                    return objDataset1;//.Tables[1];
                }
            }
        }
    }
}
