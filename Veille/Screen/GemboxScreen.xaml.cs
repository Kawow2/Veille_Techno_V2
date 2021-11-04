using System.Windows;
using System.Windows.Forms;

namespace Veille.Screen
{
    /// <summary>
    /// Logique d'interaction pour GemboxScreen.xaml
    /// </summary>
    public partial class GemboxScreen : Window
    {
        public Framework.Gembox gembox { get; set; }
        public string Filename { get; set; }

        public GemboxScreen()
        {
            InitializeComponent();
            gembox = new Framework.Gembox();
        }

        private void OpenFile_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|CSV files (*.csv)|*.csv";
            var result = openFileDialog.ShowDialog();
            if (((int)result) == 1)
            {
                if (openFileDialog.SafeFileName == this.Filename)
                    return;
                var ext = openFileDialog.SafeFileName.Split('.')[1];
                if (ext == "csv")
                {
                    //loadOptions = new LoadOptions(LoadFormat.Csv);
                }
                Filename = openFileDialog.FileName;
                var time = gembox.OpenFile(Filename);
                _Time.Content = time;
                _FileName.Content = openFileDialog.SafeFileName;
            }
        }

        private void WriteFile_Click(object sender, RoutedEventArgs e)
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
                this._Time.Content = gembox.WriteFile(filename);
            }
        }
    }
}
