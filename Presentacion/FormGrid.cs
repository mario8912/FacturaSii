using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Presentacion
{
    public partial class FormGrid : Form
    {
        private DataTable _tabla;
        public FormGrid(DataTable tabla)
        {
            InitializeComponent();
            _tabla = tabla;
        }

        private void FormGrid_Load(object sender, EventArgs e)
        {
            dataGridView1.DataSource = _tabla;
        }
    }
}
