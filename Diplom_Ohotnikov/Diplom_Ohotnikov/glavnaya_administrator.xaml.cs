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
    /// Логика взаимодействия для glavnaya_administrator.xaml
    /// </summary>
    public partial class glavnaya_administrator : Window
    {
        public glavnaya_administrator()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            adm_uslugi form = new adm_uslugi();
            form.Show();
            this.Close();
        }

        private void button2_Click(object sender, RoutedEventArgs e)
        {
            adm_klient form = new adm_klient();
            form.Show();
            this.Close();
        }

        // Функция "Выйти"
        private void button3_Click(object sender, RoutedEventArgs e)
        {
            MainWindow form = new MainWindow();
            form.Show();
            this.Close();
        }
    }
}
