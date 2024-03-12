﻿using G = Entidades.utils.Global;   
using System.Windows.Forms;

namespace Presentacion
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, System.EventArgs e)
        {
            DialogoSeleccionExcel();
        }

        private void DialogoSeleccionExcel()
        {
            OpenFileDialog openFile = new OpenFileDialog
            {
                InitialDirectory = @"E:\mipc\escritorio\FacturaSii\data",
                Filter = "Excel Files|*.xlsx",
                Title = "Selecciona un archivo"
            };

            if (openFile.ShowDialog() == DialogResult.OK && openFile.CheckFileExists == true)
                G.ExcelFile = openFile.FileName;
        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            Negocio.NegocioXml.CrearXml();
        }
    }
}
