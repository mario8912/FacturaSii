using Datos.Excel;
using Datos.XML;
using System;
using H = Entidades.utils.Helper;
namespace Negocio
{
    public class CrearXML
    {
        private readonly ConstructorXML _constructor = new ConstructorXML();
        private ExcelReader _excelReader;

        public CrearXML()
        {
            H.SetHora();
        }

        public void TryCrearXml()
        {
            try
            {
                CrearXml();
            }
            catch (Exception ex)
            {
                new Exception("Error al intentar crear la estructura XML:" 
                    + $"{Environment.NewLine}{ex.Message}");
            }   
        }

        public void CrearXml()
        {
            _excelReader = new ExcelReader();

            _constructor.EstructuraXML()?.EstructuraCabeceraXML()?.TryEstructuraFacturaXML(_excelReader.GetDiccionario());
            _constructor.TryGuardarXML();
        }

        //try catchs
    }
}
