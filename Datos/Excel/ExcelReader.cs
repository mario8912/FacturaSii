using G = Entidades.utils.Global;   
using Entidades.utils;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Data;

namespace Datos.Excel
{
    public class ExcelReader
    {
        private static Dictionary<int, dynamic> _diccionarioValores;
        private readonly DataTable _dataTable;
        private readonly Helper _listas;

        public ExcelReader()
        {
            _listas = new Helper();
            _dataTable = LeerExcelRs();
        }

        private DataTable LeerExcelRs()
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

                    return objDataset1.Tables[0];
                }
            }
        }

        public IEnumerable<Dictionary<int, dynamic>> LeerExcel()
        {
            _diccionarioValores = _listas.GetDiccionarioColumnasExcel();
            var tempDic = new Dictionary<int, dynamic>(_diccionarioValores);

            foreach (DataRow fila in _dataTable.AsEnumerable())
            {
                var tempDicForLoop = new Dictionary<int, dynamic>(_diccionarioValores);
                foreach (KeyValuePair<int, dynamic> itemDiccionario in tempDicForLoop)
                    if (fila[itemDiccionario.Key] != null)
                        tempDic[itemDiccionario.Key] = fila[itemDiccionario.Key].ToString();

                _diccionarioValores = tempDic;
                yield return _diccionarioValores;
            }
        }
    }
}
