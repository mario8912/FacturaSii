using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Datos.Excel
{
    public interface IMiProgressBar
    {
        int AumentarProgreso(int val);
        int TotalProgreso(int val);
    }
}
