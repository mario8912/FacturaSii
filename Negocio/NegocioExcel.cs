using Datos.Excel;

namespace Negocio
{
    public class NegocioExcel
    {
        public static void LeerExcel(string file)
        {
            ExcelReader er = new ExcelReader();
            er.LeerExcel(file);
        }
    }
}
