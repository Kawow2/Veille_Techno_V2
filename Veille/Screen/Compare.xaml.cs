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
                //spire
                AsposeAnalyse.Text= aspose.OpenFile(openFileDialog.FileName, loadOptions).ToString();
                SpireAnalyse.Text = spire.OpenFile(openFileDialog.FileName).ToString();
                GemboxAnalyse.Text = gembox.OpenFile(openFileDialog.FileName).ToString();
                Filename = openFileDialog.SafeFileName;
                _FileName.Content = Filename;
            }
        }

        private void WriteFile_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(this.Filename))
            {
                System.Windows.MessageBox.Show("Vous n'avez pas de chargé de fichier");
                return;
            }
            var folder = new FolderBrowserDialog();
            var result = folder.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                string filename = folder.SelectedPath + "\\" + this._FileName.Content;
                var indexLastBackSlash = filename.LastIndexOf("\\");
                var indexLastPoint= filename.LastIndexOf(".");
                var path = filename.Substring(0, indexLastBackSlash);
                var name = filename.Substring(indexLastBackSlash, indexLastPoint - indexLastBackSlash);
                var ext = filename.Substring(indexLastPoint, filename.Length - indexLastPoint);
                AsposeAnalyse.Text = aspose.WriteFile(path + name +"_aspose" + ext, SaveFormat.Auto).ToString();
                SpireAnalyse.Text = spire.WriteFile(path + name+ "_spire" + ext).ToString();
                GemboxAnalyse.Text = gembox.WriteFile(path + name + "_gembox" + ext).ToString();
            }
        }

        private void CreateChart_Click(object sender, RoutedEventArgs e)
        {
            AsposeAnalyse.Text=aspose.CreateChart().ToString();
            SpireAnalyse.Text = gembox.CreateChart().ToString();
            GemboxAnalyse.Text = spire.CreateChart().ToString();
        }
    }
}
