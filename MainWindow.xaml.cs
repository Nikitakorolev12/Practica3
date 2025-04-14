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

namespace Practica3
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private APIInteraction appinteraction;
        public MainWindow()
        {
            InitializeComponent();
            appinteraction = new APIInteraction();
        }

        private void ButtonGetFullName_Click(object sender, RoutedEventArgs e)
        {
            TextBlockFullName.Text = appinteraction.GetFullName();
        }

        private void ButtonSendResult_Click(object sender, RoutedEventArgs e)
        {
            TextBlockResult.Text = appinteraction.FillDocument();
        }
    }
}
