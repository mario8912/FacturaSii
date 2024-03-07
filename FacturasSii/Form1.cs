using FacturasSii.Utils;
using System.Windows.Forms;

namespace FacturasSii
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, System.EventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog
            {
                InitialDirectory = @"\\SERVER2017\datos\hacienda\ROSELL",
                Filter = "Excel Files|*.xlsx",
                Title = "Selecciona un archivo"
            };

            if (openFile.ShowDialog() == DialogResult.OK && openFile.CheckFileExists == true)
            {
                ExcelReader.ReadExcel(openFile.FileName);
            }
        }
    }
}