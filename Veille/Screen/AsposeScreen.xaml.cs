using Aspose.Cells;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;
using Veille.Framework;

namespace Veille.Screen
{
    /// <summary>
    /// Logique d'interaction pour AsposeScreen.xaml
    /// </summary>
    /// 

    public partial class AsposeScreen : Window
    {
        public string Filename { get; set; }
        public Framework.Aspose aspose { get; set; }

        public AsposeScreen()
        {
            InitializeComponent();
            aspose = new Framework.Aspose();
        }

        private void OpenFile(object sender, RoutedEventArgs e)
        {

            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|CSV files (*.csv)|*.csv";
            var result = openFileDialog.ShowDialog();
            if (((int)result) == 1)
            {
                Mouse.SetCursor(System.Windows.Input.Cursors.Wait);
                if (openFileDialog.FileName == this.Filename)
                {
                    System.Windows.MessageBox.Show("Ce fichier est déjà chargé");
                    return;
                }
                LoadOptions loadOptions = null;
                var ext = openFileDialog.SafeFileName.Split('.')[1];
                if (ext == "csv")
                {
                    loadOptions = new LoadOptions(LoadFormat.Csv);
                }
                Filename = openFileDialog.FileName;
                var time = aspose.OpenFile(Filename, loadOptions);
                this._Time.Content = time;
                this._FileName.Content = openFileDialog.SafeFileName;
                Mouse.SetCursor(System.Windows.Input.Cursors.Arrow);
            }
        }

        private void WriteFile(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(this.Filename))
            {
                System.Windows.MessageBox.Show("Vous n'avez pas de charger de fichier");
                return;
            }
            var folder = new FolderBrowserDialog();
            var result = folder.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                string filename = folder.SelectedPath + "\\" + this._FileName.Content;
                this._Time.Content = aspose.WriteFile(filename, SaveFormat.Auto);
            }
        }
    }
}
