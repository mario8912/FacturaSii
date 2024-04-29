using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace ManejoDeCoches
{
    public partial class MainForm : Form
    {
        private List<Coche> coches = new List<Coche>();

        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            RefrescarDataGridView();
        }

        private void RefrescarDataGridView()
        {
            dataGridViewCoches.DataSource = null;
            dataGridViewCoches.DataSource = coches;
        }

        private void buttonAgregar_Click(object sender, EventArgs e)
        {
            Coche nuevoCoche = new Coche
            {
                Marca = textBoxMarca.Text,
                Modelo = textBoxModelo.Text,
                Año = Convert.ToInt32(textBoxAño.Text)
            };
            coches.Add(nuevoCoche);
            RefrescarDataGridView();
            LimpiarCampos();
        }

        private void buttonEditar_Click(object sender, EventArgs e)
        {
            if (dataGridViewCoches.SelectedRows.Count > 0)
            {
                int indiceSeleccionado = dataGridViewCoches.SelectedRows[0].Index;
                coches[indiceSeleccionado].Marca = textBoxMarca.Text;
                coches[indiceSeleccionado].Modelo = textBoxModelo.Text;
                coches[indiceSeleccionado].Año = Convert.ToInt32(textBoxAño.Text);
                RefrescarDataGridView();
                LimpiarCampos();
            }
        }

        private void buttonEliminar_Click(object sender, EventArgs e)
        {
            if (dataGridViewCoches.SelectedRows.Count > 0)
            {
                int indiceSeleccionado = dataGridViewCoches.SelectedRows[0].Index;
                coches.RemoveAt(indiceSeleccionado);
                RefrescarDataGridView();
                LimpiarCampos();
            }
        }

        private void LimpiarCampos()
        {
            textBoxMarca.Clear();
            textBoxModelo.Clear();
            textBoxAño.Clear();
        }
    }

    public class Coche
    {
        public string Marca { get; set; }
        public string Modelo { get; set; }
        public int Año { get; set; }
    }
}
