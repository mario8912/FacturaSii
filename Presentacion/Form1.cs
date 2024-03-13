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
        private EventoProgreso _eventoProgreso;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            _eventoProgreso = new EventoProgreso();
            _eventoProgreso.ProgresoCambiado += ProgresoCambiado;

            btnCrearXml.Enabled = false;
            progressBar1.Visible = false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SelectorDeArchivo();
        }

        private void SelectorDeArchivo()
        {
            OpenFileDialog openFile = new OpenFileDialog
            {
                InitialDirectory = @"E:\mipc\escritorio\FacturaSii\data",
                Filter = "Excel Files|*.xlsx",
                Title = "Selecciona un archivo"
            };

            if (openFile.ShowDialog() == DialogResult.OK && openFile.CheckFileExists == true)
            {
                G.ExcelFile = openFile.FileName;
                btnCrearXml.Enabled = true;
                AgregarRutaTextbox();
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

            var task = new Task(() => { negocioXml.CrearXml(_eventoProgreso); });
            task.Start();   
            
            //TestProgressBar();
            ProgressBarStatus();

            await task;
            MensajeXMLCreado();
            LimpiarRecursos();
        }

        private void TestProgressBar()
        {
            _eventoProgreso.ValorMaximoBarraProgreso = 150;
            for (int i = 0; i < 100; i++)
            {
                _eventoProgreso.AumentarProgreso();
            }
        }

        private void ProgressBarStatus()
        {
            progressBar1.Visible = true;
            progressBar1.Minimum = 0;
            progressBar1.Maximum = _eventoProgreso.ValorMaximoBarraProgreso;
        }

        private void ProgresoCambiado(object sender, int aumento)
        {

            if (InvokeRequired)
            {
                Invoke(new Action(() => { SetMaxProgressBar(); }));
                Invoke(new Action(() => progressBar1.Value = aumento));
            }
                
            else
                progressBar1.Value = aumento;
        }

        private void SetMaxProgressBar()
        {
            if(progressBar1.Maximum != _eventoProgreso.ValorMaximoBarraProgreso)
                progressBar1.Maximum = _eventoProgreso.ValorMaximoBarraProgreso-1;
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
