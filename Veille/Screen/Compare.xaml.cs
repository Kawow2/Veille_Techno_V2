using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Veille.Screen
{
    /// <summary>
    /// Logique d'interaction pour Compare.xaml
    /// </summary>
    public partial class Compare : Window
    {
        public Framework.Spire spire { get; set; }
        public Framework.Gembox gembox { get; set; }
        public Framework.Aspose aspose { get; set; }
        public string Filename { get; set; }


        public Compare()
        {
            InitializeComponent();
            spire = new Framework.Spire();
            gembox = new Framework.Gembox();
            aspose = new Framework.Aspose();

        }



        private void OpenFile_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|CSV files (*.csv)|*.csv";
            var result = openFileDialog.ShowDialog();
            if (((int)result) == 1)
            {
                if (openFileDialog.FileName == this.Filename)
                    return;
                //aspose
                LoadOptions loadOptions = null;
                var ext = openFileDialog.SafeFileName.Split('.')[1];
                if (ext == "csv")
                {
                    loadOptions = new LoadOptions(LoadFormat.Csv);
                }
                Filename = openFileDialog.FileName;            
                //spire
                AsposeAnalyse.Text= aspose.OpenFile(Filename, loadOptions).ToString();
                SpireAnalyse.Text = spire.OpenFile(openFileDialog.FileName).ToString();
                GemboxAnalyse.Text = gembox.OpenFile(openFileDialog.FileName).ToString();
            }
        }

        private void WriteFile_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
