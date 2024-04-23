using System;
using System.Data;
using System.Windows.Forms;

namespace Presentacion
{
    public partial class FormGrid : Form
    {
        private readonly DataTable _tabla;
        public FormGrid(DataTable tabla)
        {
            InitializeComponent();
            _tabla = tabla;
        }

        private void FormGrid_Load(object sender, EventArgs e)
        {
            dataGridView1.DataSource = _tabla;
        }

        private void FormGrid_FormClosing(object sender, FormClosingEventArgs e)
        {
            Dispose();
        }
    }
}
