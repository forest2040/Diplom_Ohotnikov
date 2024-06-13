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
using System.Data;
using System.Data.SqlClient;

namespace Diplom_Ohotnikov
{
    /// <summary>
    /// Логика взаимодействия для status.xaml
    /// </summary>
    public partial class status : Window
    {
        SqlCommand cmd;
        SqlConnection con;
        SqlDataAdapter da;
        DataSet ds;
        public status()
        {
            InitializeComponent();
            GetList();
        }
        void GetList()
        {
            con = new SqlConnection(@"data source = stud-mssql.sttec.yar.ru,38325; initial catalog = user57_db; user id = user57_db; password = user57");
            con.Open();
            da = new SqlDataAdapter("SELECT * FROM Diplom_Status", con);
            ds = new DataSet();
            da.Fill(ds, "Diplom_Status");
            DataView dview = new DataView(ds.Tables["Diplom_Status"]);
            dataGridView1.ItemsSource = dview;
            con.Close();
        }
        void GetList2()
        {
            con = new SqlConnection(@"data source = stud-mssql.sttec.yar.ru,38325; initial catalog = user57_db; user id = user57_db; password = user57");
            con.Open();
            da = new SqlDataAdapter("SELECT * FROM Diplom_Sotrudnik", con);
            ds = new DataSet();
            da.Fill(ds, "Diplom_Sotrudnik");
            con.Close();
        }
        // Функция "Выйти"
        private void button1_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
        // Функция "Добавить"
        private void button2_Click(object sender, RoutedEventArgs e)
        {
            if (textBox2.Text != "")
            {
                cmd = new SqlCommand();
                cmd.Connection = con;
                con.Open();
                cmd.CommandText = "INSERT INTO Diplom_Status (status) values (" + textBox2.Text + ")";
                cmd.ExecuteNonQuery();
                con.Close();
                GetList();
                MessageBox.Show(string.Format("Успешное добавление данных в таблицу"), "Сообщение");
            }
            else
            {
                MessageBox.Show("Одно из текстовых полей не заполнено", "Ошибка");
            }
        }
        // Функция "Изменить"
        private void button3_Click(object sender, RoutedEventArgs e)
        {
            if (textBox1.Text != "")
            {
                cmd = new SqlCommand();
                cmd.Connection = con;
                con.Open();
                cmd.CommandText = "UPDATE Diplom_Status SET status = '" + textBox2.Text + "' where kod_statusa = '" + textBox1.Text + "'";
                cmd.ExecuteNonQuery();
                con.Close();
                GetList();
                MessageBox.Show(string.Format("Успешное изменение данных в таблице"), "Сообщение");
            }
            else
            {
                MessageBox.Show("Текстовое поле 'Код статуса' не заполнено", "Ошибка");
            }
        }
        // Функция "Удалить"
        private void button4_Click(object sender, RoutedEventArgs e)
        {
            if (textBox1.Text != "")
            {
                cmd = new SqlCommand();
                cmd.Connection = con;
                con.Open();
                cmd.CommandText = "DELETE FROM Diplom_Status WHERE kod_statusa = '" + textBox1.Text + "'";
                cmd.ExecuteNonQuery();
                con.Close();
                GetList();
                MessageBox.Show(string.Format("Успешное удаление данных в таблице"), "Сообщение");
                textBox1.Clear();
                textBox2.Clear();
            }
            else
            {
                MessageBox.Show("Текстовое поле 'Код статуса' не заполнено", "Ошибка");
            }
        }
        // Функция "Изменить статус сотрудника"
        private void button5_Click(object sender, RoutedEventArgs e)
        {
            if (textBox3.Text != "")
            {
                cmd = new SqlCommand();
                cmd.Connection = con;
                con.Open();
                cmd.CommandText = "UPDATE Diplom_Sotrudnik SET kod_statusa = '" + comboBox1.Text + "' where kod_sotrudnika = '" + textBox3.Text + "'";
                cmd.ExecuteNonQuery();
                con.Close();
                GetList2();
                MessageBox.Show(string.Format("Успешное изменение данных в таблице 'Сотрудник'"), "Сообщение");
            }
            else
            {
                MessageBox.Show("Текстовое поле 'Код сотрудника' не заполнено", "Ошибка");
            }
        }
    }
}
