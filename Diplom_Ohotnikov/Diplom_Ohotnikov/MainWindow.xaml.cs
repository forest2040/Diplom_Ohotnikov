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
using System.Data;
using System.Data.SqlClient;

namespace Diplom_Ohotnikov
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            string connect = "data source = stud-mssql.sttec.yar.ru,38325; initial catalog = user57_db; user id = user57_db; password = user57";
            string sql = "SELECT * FROM Diplom_Sotrudnik WHERE login = '" + textBox1.Text + "' and parol = '" + textBox2.Text + "'";
            SqlConnection sqlConnection = new SqlConnection(connect);
            SqlCommand sqlCommand = new SqlCommand(sql, sqlConnection);
            sqlConnection.Open();
            SqlDataReader dr = sqlCommand.ExecuteReader();
            string login = "null";
            string parol = "null";
            string rol = "null";
            while (dr.Read())
            {
                login = dr.GetString(11);
                parol = dr.GetString(12);
                rol = dr.GetString(5);
            }
            sqlConnection.Close();
            if ((login == "null") || (parol == "null"))
            {
                MessageBox.Show(string.Format("Неправильный логин или пароль"), "Ошибка");
            }
            else
            {
                if (rol == "Директор")
                {
                    MessageBox.Show(string.Format("Вы успешно авторизовались в роли 'Директор'"), "Сообщение");
                    glavnaya_director form = new glavnaya_director();
                    form.Show();
                    this.Close();
                }
                else if (rol == "Администратор")
                {
                    MessageBox.Show(string.Format("Вы успешно авторизовались в роли 'Администратор'"), "Сообщение");
                    glavnaya_administrator form = new glavnaya_administrator();
                    form.Show();
                    this.Close();
                }
                else if (rol == "Врач")
                {
                    MessageBox.Show(string.Format("Вы успешно авторизовались в роли 'Врач'"), "Сообщение");
                    glavnaya_vrach form = new glavnaya_vrach();
                    form.Show();
                    this.Close();
                }
            }
        }
    }
}