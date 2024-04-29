using G = Entidades.utils.Global;
using Entidades.utils;
using Negocio;
using Negocio.NegocioXML;
using System;
using System.Data;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Presentacion
{
    public partial class Form1 : Form
    {
        private static OpenFileDialog _openFileDialog;
        private readonly CrearXML _negocioCrearXML;
        
        public Form1()
        {
            BringToFront();
            InitializeComponent();
            _negocioCrearXML = new CrearXML();
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
                InitialDirectory = G.RutaDirectorioData,
                Filter = "Excel Files|*.xlsx",
                Title = "Selecciona un archivo"
            };

            return this;
        }

        private readonly Func<bool> ComprobarArchivoSeleccionadoExiste = () => (_openFileDialog.ShowDialog() == DialogResult.OK && _openFileDialog.CheckFileExists);

        private async void btnCrearXml_Click(object sender, EventArgs e)
        {
            FormatearControles();

            await TaskCrearXml();

            MensajeXMLCreado();

            ProcesadoEnvioRespuesta();

            LimpiarRecursos();
        }

        private void FormatearControles()
        {
            textBox1.Text = string.Empty;
            btnCrearXml.Enabled = false;
        }

        private Task TaskCrearXml()
        {
            Task task = Task.Run(() =>
            {
                _negocioCrearXML.CrearXml();
            });

            return task;
        }

        private void MensajeXMLCreado()
        {
            DialogResult result = MessageBox.Show(
                "XML creado\n¿Desea visualizarlo?",
                "Aviso",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
                );

            if (result is DialogResult.Yes)
                Process.Start(Global.RutaGuardarXmlEnvio);
        }

        private async void ProcesadoEnvioRespuesta()
        {
            await TaskEnvio(); 
            new FormGrid(new RespuestaXML().ProcesarRespuesta()).Show();
        }

        private Task TaskEnvio()
        {
            Task task = Task.Run(() =>
            {
                new Enviar().Envio();
            });

            return task;
        }

        private void LimpiarRecursos()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            LimpiarRecursos();
            CerrarAplicacion();
        }

        private void CerrarAplicacion()
        {
            Dispose();
            Close();
        }

        private void Form1_Activated(object sender, EventArgs e)
        {
            btnSeleccionArchivo.Focus();
        }
    }
}
