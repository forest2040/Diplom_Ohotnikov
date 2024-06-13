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
    /// Логика взаимодействия для vr_okazannye_uslugi.xaml
    /// </summary>
    public partial class vr_okazannye_uslugi
    {
        private string connectionString = "data source = stud-mssql.sttec.yar.ru,38325; initial catalog = user57_db; user id = user57_db; password = user57";
        SqlCommand cmd;
        SqlConnection con;
        SqlDataAdapter da;
        DataSet ds;
        public vr_okazannye_uslugi()
        {
            InitializeComponent();
            GetList();
            LoadData();
        }
        void GetList()
        {
            con = new SqlConnection(@"data source = stud-mssql.sttec.yar.ru,38325; initial catalog = user57_db; user id = user57_db; password = user57");
            con.Open();
            da = new SqlDataAdapter("SELECT * FROM Diplom_OkazannyeUslugi", con);
            ds = new DataSet();
            da.Fill(ds, "Diplom_OkazannyeUslugi");
            DataView dview = new DataView(ds.Tables["Diplom_OkazannyeUslugi"]);
            dataGridView1.ItemsSource = dview;
            con.Close();
        }
        private void LoadData()
        {
            string query = "SELECT * FROM Diplom_OkazannyeUslugi";
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
            if ((textBox1.Text != "") && (textBox2.Text != "") && (textBox3.Text != "") && (textBox4.Text != ""))
            {
                cmd = new SqlCommand();
                cmd.Connection = con;
                con.Open();
                cmd.CommandText = "INSERT INTO Diplom_OkazannyeUslugi (naimenovanie,stoimost,data_okazania,kod_priema) values (" + textBox1.Text + ",'" + textBox2.Text + "','" + textBox3.Text + "','" + textBox4.Text + "')";
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
                cmd.CommandText = "UPDATE Diplom_OkazannyeUslugi SET naimenovanie = '" + textBox1.Text + "',stoimost='" + textBox2.Text + "',data_okazania='" + textBox3.Text + "',kod_priema='"
                + textBox4.Text + "' where kod_okazannoy_uslugi = '" + textBox6.Text + "'";
                cmd.ExecuteNonQuery();
                con.Close();
                GetList();
                MessageBox.Show(string.Format("Успешное изменение данных в таблице"), "Сообщение");
            }
            else
            {
                MessageBox.Show("Текстовое поле 'Код оказанной услуги' не заполнено", "Ошибка");
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
                cmd.CommandText = "DELETE FROM Diplom_OkazannyeUslugi WHERE kod_okazannoy_uslugi = '" + textBox6.Text + "'";
                cmd.ExecuteNonQuery();
                con.Close();
                GetList();
                MessageBox.Show(string.Format("Успешное удаление данных в таблице"), "Сообщение");
                textBox1.Clear();
                textBox2.Clear();
                textBox3.Clear();
                textBox4.Clear();
                textBox6.Clear();
            }
            else
            {
                MessageBox.Show("Текстовое поле 'Код оказанной услуги' не заполнено", "Ошибка");
            }
        }
        // Функция "Вывод в Excel"
        private void button5_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            excel.Visible = true;
            Workbook workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);
            Worksheet sheet1 = (Worksheet)workbook.Sheets[1];

            for (int j = 0; j < dataGridView1.Columns.Count; j++)
            {
                Range myRange = (Range)sheet1.Cells[1, j + 1];
                sheet1.Cells[1, j + 1].Font.Bold = true;
                sheet1.Columns[j + 1].ColumnWidth = 15;
                myRange.Value2 = dataGridView1.Columns[j].Header;
            }
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Items.Count; j++)
                {
                    TextBlock b = dataGridView1.Columns[i].GetCellContent(dataGridView1.Items[j]) as TextBlock;
                    Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[j + 2, i + 1];
                    myRange.Value2 = b.Text;
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
            vr_okazannye_uslugi form = new vr_okazannye_uslugi();
            form.Show();
            this.Close();
        }
        // Функция "Сортировка по возрастанию"
        private void button8_Click(object sender, RoutedEventArgs e)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                SqlCommand command = new SqlCommand("SELECT * FROM Diplom_OkazannyeUslugi ORDER BY naimenovanie ASC", connection);
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
                SqlCommand command = new SqlCommand("SELECT * FROM Diplom_OkazannyeUslugi ORDER BY naimenovanie DESC", connection);
                SqlDataReader reader = command.ExecuteReader();
                System.Data.DataTable table = new System.Data.DataTable();
                table.Load(reader);
                dataGridView1.ItemsSource = table.DefaultView;
                reader.Close();
            }
        }
        // Функция "Выйти"
        private void button1_Click(object sender, RoutedEventArgs e)
        {
            glavnaya_vrach form = new glavnaya_vrach();
            form.Show();
            this.Close();
        }
    }
}
