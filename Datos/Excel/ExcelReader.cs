using G = Entidades.utils.Global;   
using Entidades.utils;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Data;
using System;
using System.Linq;

namespace Datos.Excel
{
    public class ExcelReader
    {
        private const string XLS_CONNECTION_STRING = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=Excel 8.0;";
        private const string XLSX_CONNECTION_STRING = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=Excel 12.0;";

        private readonly Dictionary<int, dynamic> _diccionarioValores;
        private readonly DataTable _dataTable;
        private readonly Helper _listas;
        public ExcelReader()
        {
            _listas = new Helper();
            _dataTable = PasarExcelADataTable();
            _diccionarioValores = _listas.GetDiccionarioColumnasExcel();
        }

        private DataTable PasarExcelADataTable()
        {
            var sConnectionString = string.Format(XLS_CONNECTION_STRING, G.ExcelFile);

            using (_ = new Excel())
            using (OleDbConnection objConn = new OleDbConnection(sConnectionString))
            {
                var objCon1 = TryOpenConnection(objConn);

                OleDbCommand objCmdSelect = new OleDbCommand("SELECT * FROM [Hoja1$]", objCon1);

                OleDbDataAdapter objAdapter1 = new OleDbDataAdapter
                {
                    SelectCommand = objCmdSelect
                };

                DataSet objDataset1 = new DataSet();
                objAdapter1.Fill(objDataset1);

                if (objCon1.State == ConnectionState.Open)
                    objCon1.Close();

                objCon1.Dispose();

                return objDataset1.Tables[0];
            }
        }
        
        private OleDbConnection TryOpenConnection(OleDbConnection objConn)
        {
            try
            {
                objConn.Open();
                return objConn;
            }
            catch (OleDbException)
            {
                return new OleDbConnection(string.Format(XLSX_CONNECTION_STRING, G.ExcelFile));
            }
            catch (Exception ex)
            {
                throw new Exception("Error al abrir la conexión con el archivo Excel.", ex);
            }
        }

        public IEnumerable<Dictionary<int, dynamic>> GetDiccionario()
        {
            foreach (DataRow fila in _dataTable.AsEnumerable())
                yield return AsignarValoresAlDiccionario(_diccionarioValores, fila);
        }

        //Se crean diccionarios temporales, que son copias del diccionario original ya que
        //al recorrer un diccionario y cambiar sus valores en tiempo de ejecución, se produce una excepción.
        private Dictionary<int, dynamic> AsignarValoresAlDiccionario(Dictionary<int, dynamic> diccionarioValores, DataRow fila)
        {
            var diccionarioValoresTemporal = new Dictionary<int, dynamic>(diccionarioValores);

            foreach (KeyValuePair<int, dynamic> itemDiccionarioColumna in diccionarioValoresTemporal)
            {
                var valor = fila[itemDiccionarioColumna.Key];

                if (valor != null)
                    if(itemDiccionarioColumna.Key == 1) 
                        valor = valor.ToString().Replace('/', '-');
                diccionarioValores[itemDiccionarioColumna.Key] = valor.ToString();
            }
            return diccionarioValores;
        }
    }
}
