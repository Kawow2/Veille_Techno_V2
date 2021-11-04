using System.Windows;
using System.Windows.Forms;
using Veille.Framework;
namespace Veille.Screen
{
    /// <summary>
    /// Logique d'interaction pour SpireScreen.xaml
    /// </summary>
    public partial class SpireScreen : Window
    {
        public Framework.Spire spire { get; set; }
        public string Filename { get; set; }
        public SpireScreen()
        {
            InitializeComponent();
            spire = new Framework.Spire();
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
                this.Filename = openFileDialog.SafeFileName;
                this._Time.Content = spire.OpenFile(openFileDialog.FileName);
                this._FileName.Content = this.Filename;
            }
        }

        private void WriteFile_Click(object sender, RoutedEventArgs e)
        {
            var folder = new FolderBrowserDialog();
            if (string.IsNullOrEmpty(this.Filename))
            {
                System.Windows.MessageBox.Show("Vous n'avez pas chargé de fichier");
                return;
            }

            var result = folder.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                var filename = folder.SelectedPath + "\\" + this._FileName.Content;
                this._Time.Content = spire.WriteFile(filename);
            }
        }
    }
}
