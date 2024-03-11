using Entidades.utils;
using Datos.Excel;
using System.Collections.Generic;
using System;

namespace Negocio
{
    public class NegocioExcel
    {
        public static void LeerExcel(string file)
        {
            ExcelReader.LeerExcel(file);
        }
    }
}
