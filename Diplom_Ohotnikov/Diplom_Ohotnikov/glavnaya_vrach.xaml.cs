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
using System.Windows.Shapes;

namespace Diplom_Ohotnikov
{
    /// <summary>
    /// Логика взаимодействия для glavnaya_vrach.xaml
    /// </summary>
    public partial class glavnaya_vrach : Window
    {
        public glavnaya_vrach()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, RoutedEventArgs e)
        {
            MainWindow form = new MainWindow();
            form.Show();
            this.Close();
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            vr_uslugi form = new vr_uslugi();
            form.Show();
            this.Close();
        }

        private void button2_Click(object sender, RoutedEventArgs e)
        {
            vr_priem form = new vr_priem();
            form.Show();
            this.Close();
        }

        private void button4_Click(object sender, RoutedEventArgs e)
        {
            vr_okazannye_uslugi form = new vr_okazannye_uslugi();
            form.Show();
            this.Close();
        }
    }
}
