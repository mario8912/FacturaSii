using System;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace Entidades.utils
{
    internal class GestorErrores
    {
        private readonly Exception _ex;
        private static dynamic _error;
        private static dynamic _valor;

        /*internal GestorErrores(Exception ex) 
        { 
            _valor = ex;
        }*/
        internal static dynamic TryParseFloat(dynamic valor)
        {
            _valor = valor;
			try
			{
                return float.Parse(valor);
            }
			catch (FormatException fEx)
			{
                _error = fEx;

                var infoEorr = FormatoInformacionErorror(); 
                
                GuardarErrorEnLog(infoEorr);
                GestionDeError();
                
                return 0;
            }
        }

        private static string FormatoInformacionErorror()
        {
            return string.Format(
                  "El dato entrante es  \"as\" {0} cuando debería ser numérico. {1}{2}{3}",
                  _valor, Environment.NewLine, _error.Message, Environment.NewLine
                  );
        }

        private static void GestionDeError()
        {
            MessageBox.Show("La aplicación se reiniciará.");
            Thread.Sleep(2000);
            Application.Restart();
        }

        private static void GuardarErrorEnLog(string log)
        {
            StreamWriter sw = new StreamWriter(Environment.CurrentDirectory + @"\log.txt", true);
            sw.WriteLine(log);
            sw.Close();
        }
    }
}
