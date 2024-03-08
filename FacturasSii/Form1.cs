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
                InitialDirectory = @"E:\mipc\escritorio\FacturasSii\data",
                Filter = "Excel Files|*.xlsx",
                Title = "Selecciona un archivo"
            };

            if (openFile.ShowDialog() == DialogResult.OK && openFile.CheckFileExists == true)
            {
                ExcelReader er = new ExcelReader();
                er.ReadExcel(openFile.FileName);
            }
        }
    }
}