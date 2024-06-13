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
    /// Логика взаимодействия для okazannye_uslugi.xaml
    /// </summary>
    public partial class okazannye_uslugi
    {
        private string connectionString = "data source = stud-mssql.sttec.yar.ru,38325; initial catalog = user57_db; user id = user57_db; password = user57";
        SqlConnection con;
        SqlDataAdapter da;
        DataSet ds;
        public okazannye_uslugi()
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
        // Функция "Выйти"
        private void button1_Click(object sender, RoutedEventArgs e)
        {
            glavnaya_director form = new glavnaya_director();
            form.Show();
            this.Close();
        }
        // Функция "Вывод в Excel"
        private void button2_Click(object sender, RoutedEventArgs e)
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
        // Функция "Рассчитать стоимость оказанных услуг"
        private void button3_Click(object sender, RoutedEventArgs e)
        {
            string connectionString = "data source = stud-mssql.sttec.yar.ru,38325; initial catalog = user57_db; user id = user57_db; password = user57";
            string query = "SELECT SUM(stoimost) FROM Diplom_OkazannyeUslugi";
            GetList();
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    connection.Open();
                    int total = (int)command.ExecuteScalar();

                    label1.Content = $"{total} рублей";
                }
            }
        }
        // Функция "Рассчитать общее количество оказанных услуг"
        private void button4_Click(object sender, RoutedEventArgs e)
        {
            string connectionString = "data source = stud-mssql.sttec.yar.ru,38325; initial catalog = user57_db; user id = user57_db; password = user57";
            string query = "SELECT COUNT(kod_okazannoy_uslugi) FROM Diplom_OkazannyeUslugi";
            GetList();
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    connection.Open();
                    int total = (int)command.ExecuteScalar();

                    label2.Content = $"{total}";
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
            okazannye_uslugi form = new okazannye_uslugi();
            form.Show();
            this.Close();
        }
        // Функция "Сортировка по возрастанию"
        private void button8_Click(object sender, RoutedEventArgs e)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                SqlCommand command = new SqlCommand("SELECT * FROM Diplom_OkazannyeUslugi ORDER BY data_okazania ASC", connection);
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
                SqlCommand command = new SqlCommand("SELECT * FROM Diplom_OkazannyeUslugi ORDER BY data_okazania DESC", connection);
                SqlDataReader reader = command.ExecuteReader();
                System.Data.DataTable table = new System.Data.DataTable();
                table.Load(reader);
                dataGridView1.ItemsSource = table.DefaultView;
                reader.Close();
            }
        }
    }
}
