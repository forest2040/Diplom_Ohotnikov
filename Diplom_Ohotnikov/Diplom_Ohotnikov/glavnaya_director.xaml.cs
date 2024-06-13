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
    /// Логика взаимодействия для glavnaya_director.xaml
    /// </summary>
    public partial class glavnaya_director : Window
    {
        public glavnaya_director()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            uslugi form = new uslugi();
            form.Show();
            this.Close();
        }

        private void button7_Click(object sender, RoutedEventArgs e)
        {
            MainWindow form = new MainWindow();
            form.Show();
            this.Close();
        }

        private void button6_Click(object sender, RoutedEventArgs e)
        {
            medcenter form = new medcenter();
            form.Show();
            this.Close();
        }

        private void button2_Click(object sender, RoutedEventArgs e)
        {
            klient form = new klient();
            form.Show();
            this.Close();
        }

        private void button3_Click(object sender, RoutedEventArgs e)
        {
            sotrudniki form = new sotrudniki();
            form.Show();
            this.Close();
        }

        private void button4_Click(object sender, RoutedEventArgs e)
        {
            okazannye_uslugi form = new okazannye_uslugi();
            form.Show();
            this.Close();
        }

        private void button5_Click(object sender, RoutedEventArgs e)
        {
            priem form = new priem();
            form.Show();
            this.Close();
        }
    }
}
