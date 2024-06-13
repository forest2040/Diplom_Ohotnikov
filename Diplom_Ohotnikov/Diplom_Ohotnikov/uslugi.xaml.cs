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
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using Window = Microsoft.Office.Interop.Excel.Window;

namespace Diplom_Ohotnikov
{
    /// <summary>
    /// Логика взаимодействия для uslugi.xaml
    /// </summary>
    public partial class uslugi
    {
        private string connectionString = "data source = stud-mssql.sttec.yar.ru,38325; initial catalog = user57_db; user id = user57_db; password = user57";
        SqlCommand cmd;
        SqlConnection con;
        SqlDataAdapter da;
        DataSet ds;
        public uslugi()
        {
            InitializeComponent();
            GetList();
            LoadData();
        }
        void GetList()
        {
            con = new SqlConnection(@"data source = stud-mssql.sttec.yar.ru,38325; initial catalog = user57_db; user id = user57_db; password = user57");
            con.Open();
            da = new SqlDataAdapter("SELECT * FROM Diplom_Usluga", con);
            ds = new DataSet();
            da.Fill(ds, "Diplom_Usluga");
            DataView dview = new DataView(ds.Tables["Diplom_Usluga"]);
            dataGridView1.ItemsSource = dview;
            con.Close();
        }
        private void LoadData()
        {
            string query = "SELECT * FROM Diplom_Usluga";
            SqlConnection connection = new SqlConnection(connectionString);
            da = new SqlDataAdapter(query, connection);
            SqlCommandBuilder commandBuilder = new SqlCommandBuilder(da);
            System.Data.DataTable table = new System.Data.DataTable();
            da.Fill(table);
            dataGridView1.ItemsSource = table.DefaultView;
        }
        // Функция "Добавить"
        private void button2_Click(object sender, RoutedEventArgs e)
        {
            if ((textBox1.Text != "") && (textBox2.Text != "") && (textBox3.Text != "") && (textBox4.Text != "") && (textBox5.Text != "") && (textBox7.Text != "") && (comboBox1.Text != ""))
            {
                cmd = new SqlCommand();
                cmd.Connection = con;
                con.Open();
                cmd.CommandText = "INSERT INTO Diplom_Usluga (naimenovanie,stoimost,kategoria_uslug,status,srok_vypolnenya,raspisanie,kod_centra) values (" + textBox1.Text + ",'" + textBox2.Text + "','" + textBox3.Text + "','" + textBox4.Text + "','" + textBox5.Text + "','" + textBox7.Text + "','" + comboBox1.Text + "')";
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
            if (textBox6.Text != "")
            {
                cmd = new SqlCommand();
                cmd.Connection = con;
                con.Open();
                cmd.CommandText = "UPDATE Diplom_Usluga SET naimenovanie = '" + textBox1.Text + "',stoimost='" + textBox2.Text + "',kategoria_uslug='" + textBox3.Text + "',status='"
                + textBox4.Text + "',srok_vypolnenya='" + textBox5.Text + "',raspisanie='" + textBox7.Text + "',kod_centra='" + comboBox1.Text + "' where kod_uslugi = '" + textBox6.Text + "'";
                cmd.ExecuteNonQuery();
                con.Close();
                GetList();
                MessageBox.Show(string.Format("Успешное изменение данных в таблице"), "Сообщение");
            }
            else
            {
                MessageBox.Show("Текстовое поле 'Код услуги' не заполнено", "Ошибка");
            }
        }
        // Функция "Удалить"
        private void button4_Click(object sender, RoutedEventArgs e)
        {
            if (textBox6.Text != "")
            {
                cmd = new SqlCommand();
                cmd.Connection = con;
                con.Open();
                cmd.CommandText = "DELETE FROM Diplom_Usluga WHERE kod_uslugi = '" + textBox6.Text + "'";
                cmd.ExecuteNonQuery();
                con.Close();
                GetList();
                MessageBox.Show(string.Format("Успешное удаление данных в таблице"), "Сообщение");
                textBox1.Clear();
                textBox2.Clear();
                textBox3.Clear();
                textBox4.Clear();
                textBox5.Clear();
                textBox6.Clear();
                textBox7.Clear();
                textBox15.Clear();
                comboBox1.Items.Clear();
            }
            else
            {
                MessageBox.Show("Текстовое поле 'Код услуги' не заполнено", "Ошибка");
            }
        }
        // Функция "Выйти"
        private void button1_Click(object sender, RoutedEventArgs e)
        {
            glavnaya_director form = new glavnaya_director();
            form.Show();
            this.Close();
        }

        private void button5_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            excel.Visible = true;
            Workbook workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);
            Worksheet sheet1 = (Worksheet)workbook.Sheets[1];

            int rowCount = dataGridView1.Items.Count;
            int columnCount = dataGridView1.Columns.Count;

            for (int j = 0; j < columnCount; j++)
            {
                Range headerRange = (Range)sheet1.Cells[1, j + 1];
                headerRange.Font.Bold = true;
                sheet1.Columns[j + 1].ColumnWidth = 15;
                headerRange.Value2 = dataGridView1.Columns[j].Header;
            }

            for (int i = 0; i < columnCount; i++)
            {
                for (int j = 0; j < rowCount; j++)
                {
                    TextBlock textBlock = dataGridView1.Columns[i].GetCellContent(dataGridView1.Items[j]) as TextBlock;
                    if (textBlock != null)
                    {
                        Range cellRange = (Range)sheet1.Cells[j + 2, i + 1];
                        cellRange.Value2 = textBlock.Text;
                    }
                }
            }
        }
        // Функция "Поиск по наименованию"
        private void button6_Click(object sender, RoutedEventArgs e)
        {
            if (textBox15.Text != "")
            {
                string filter = textBox15.Text;
                DataView view = dataGridView1.ItemsSource as DataView;
                if (view != null)
                {
                    view.RowFilter = $"naimenovanie LIKE '%{filter}%'";
                }
            }
            else
            {
                MessageBox.Show("Текстовое поле 'Поиск по наименованию' не заполнено", "Ошибка");
            }
        }
        // Функция "Очистить"
        private void button7_Click(object sender, RoutedEventArgs e)
        {
            textBox15.Clear();
            uslugi form = new uslugi();
            form.Show();
            this.Close();
        }
        // Функция "Сортировка по возрастанию"
        private void button8_Click(object sender, RoutedEventArgs e)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                SqlCommand command = new SqlCommand("SELECT * FROM Diplom_Usluga ORDER BY stoimost ASC", connection);
                SqlDataReader reader = command.ExecuteReader();
                System.Data.DataTable table = new System.Data.DataTable();
                table.Load(reader);
                dataGridView1.ItemsSource = table.DefaultView;
                reader.Close();
            }
        }
        // Функция "Сортировка по убыванию"
        private void button9_Click(object sender, RoutedEventArgs e)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                SqlCommand command = new SqlCommand("SELECT * FROM Diplom_Usluga ORDER BY stoimost DESC", connection);
                SqlDataReader reader = command.ExecuteReader();
                System.Data.DataTable table = new System.Data.DataTable();
                table.Load(reader);
                dataGridView1.ItemsSource = table.DefaultView;
                reader.Close();
            }
        }
    }
}