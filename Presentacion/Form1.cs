using G = Entidades.utils.Global;
using Entidades.utils;
using System.Windows.Forms;
using System;
using System.Diagnostics;
using System.Threading.Tasks;
using Negocio;

namespace Presentacion
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            btnCrearXml.Enabled = false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SelectorDeArchivo();
        }

        private void SelectorDeArchivo()
        {
            OpenFileDialog dialogoElegirExcel = new OpenFileDialog
            {
                InitialDirectory = @"E:\mipc\escritorio\FacturaSii\data",
                Filter = "Excel Files|*.xlsx",
                Title = "Selecciona un archivo"
            };

            if (dialogoElegirExcel.ShowDialog() == DialogResult.OK && dialogoElegirExcel.CheckFileExists == true)
            {
                G.ExcelFile = dialogoElegirExcel.FileName;
                btnCrearXml.Enabled = true;
                AgregarRutaTextbox();

                dialogoElegirExcel.Dispose();
            }
        }

        private void AgregarRutaTextbox()
        {
            textBox1.Text = G.ExcelFile;
        }

        private async void btnCrearXml_Click(object sender, EventArgs e)
        {
            NegocioXml negocioXml = new NegocioXml();

            textBox1.Text = string.Empty;
            btnCrearXml.Enabled = false;

            var task = new Task(() => { negocioXml.CrearXml(); });
            task.Start();

            await task;
            MensajeXMLCreado();
            LimpiarRecursos();
        }

        private void MensajeXMLCreado()
        {
            DialogResult result = MessageBox.Show("XML creado\n¿Desea visualizarlo?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result is DialogResult.Yes)
                Process.Start(G.RutaGuardarXml);
        }

        private void LimpiarRecursos()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Dispose();
            Close();
        }   
    }
}
