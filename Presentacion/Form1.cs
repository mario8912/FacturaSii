using Entidades.utils;
using Negocio;
using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Presentacion
{
    public partial class Form1 : Form
    {
        private static OpenFileDialog _openFileDialog;
        public Form1()
        {
            InitializeComponent();
        }

        private void botonSelecionArchivo_Click(object sender, EventArgs e)
        {
            if (MostrarSelectorDeArchivo().ComprobarArchivoSeleccionadoExiste())
            {
                textBox1.Text = Global.ExcelFile = _openFileDialog.FileName;
                btnCrearXml.Enabled = true;
            }

            btnCrearXml.Focus();
            _openFileDialog.Dispose();
        }

        private Form1 MostrarSelectorDeArchivo()
        {
            _openFileDialog = new OpenFileDialog
            {
                InitialDirectory = @"E:\mipc\escritorio\FacturaSii\data",
                Filter = "Excel Files|*.xlsx",
                Title = "Selecciona un archivo"
            };

            return this;
        }

        private readonly Func<bool> ComprobarArchivoSeleccionadoExiste = () => (_openFileDialog.ShowDialog() == DialogResult.OK && _openFileDialog.CheckFileExists);

        private async void btnCrearXml_Click(object sender, EventArgs e)
        {
            NegocioXml negocioXml = new NegocioXml();

            textBox1.Text = string.Empty;
            btnCrearXml.Enabled = false;
            
            negocioXml.CrearXml();
            /*var task = new Task(() => { negocioXml.CrearXml(); });
            task.Start();

            await task;*/
            MensajeXMLCreado();
            LimpiarRecursos();
        }

        private void MensajeXMLCreado()
        {
            DialogResult result = MessageBox.Show("XML creado\n¿Desea visualizarlo?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result is DialogResult.Yes)
                Process.Start(Global.RutaGuardarXml);
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
