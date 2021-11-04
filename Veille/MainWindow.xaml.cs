using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Veille.Screen;

namespace Veille
{
    /// <summary>
    /// Logique d'interaction pour MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void OpenAspose(object sender, RoutedEventArgs e)
        {
            var a = new AsposeScreen();
            a.Show();
        }

        private void OpenSpire(object sender, RoutedEventArgs e)
        {
            var a = new SpireScreen();
            a.Show();
        }

        private void OpenGembox(object sender, RoutedEventArgs e)
        {
            var a = new GemboxScreen();
            a.Show();
        }

        private void OpenCompare_Click(object sender, RoutedEventArgs e)
        {
            var a = new Compare();
            a.Show();
        }
    }
}
