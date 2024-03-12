using G = Entidades.utils.Global;   
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
            btnCrearXml.Enabled = false;
            progressBar1.Visible = false;
        }

        private void button2_Click(object sender, System.EventArgs e)
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

        private void btnCrearXml_Click(object sender, System.EventArgs e)
        {
            Negocio.NegocioXml.CrearXml();
            MessageBox.Show("Archivo creado con éxito");
            textBox1.Text = string.Empty;
            btnCrearXml.Enabled = false;
            progressBar1.Visible = true;
        }
    }
}
