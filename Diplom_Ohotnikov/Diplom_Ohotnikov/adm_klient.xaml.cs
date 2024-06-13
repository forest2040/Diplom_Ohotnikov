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
    /// Логика взаимодействия для adm_klient.xaml
    /// </summary>
    public partial class adm_klient
    {
        private string connectionString = "data source = stud-mssql.sttec.yar.ru,38325; initial catalog = user57_db; user id = user57_db; password = user57";
        SqlCommand cmd;
        SqlConnection con;
        SqlDataAdapter da;
        DataSet ds;
        public adm_klient()
        {
            InitializeComponent();
            GetList();
            LoadData();
        }
        void GetList()
        {
            con = new SqlConnection(@"data source = stud-mssql.sttec.yar.ru,38325; initial catalog = user57_db; user id = user57_db; password = user57");
            con.Open();
            da = new SqlDataAdapter("SELECT * FROM Diplom_Klient", con);
            ds = new DataSet();
            da.Fill(ds, "Diplom_Klient");
            DataView dview = new DataView(ds.Tables["Diplom_Klient"]);
            dataGridView1.ItemsSource = dview;
            con.Close();
        }
        private void LoadData()
        {
            string query = "SELECT * FROM Diplom_Klient";
            SqlConnection connection = new SqlConnection(connectionString);
            da = new SqlDataAdapter(query, connection);
            SqlCommandBuilder commandBuilder = new SqlCommandBuilder(da);
            System.Data.DataTable table = new System.Data.DataTable();
            da.Fill(table);
            dataGridView1.ItemsSource = table.DefaultView;
        }
        // Функция "Изменить"
        private void button3_Click(object sender, RoutedEventArgs e)
        {
            if (textBox11.Text != "")
            {
                cmd = new SqlCommand();
                cmd.Connection = con;
                con.Open();
                cmd.CommandText = "UPDATE Diplom_Klient SET familiya = '" + textBox1.Text + "',imya='" + textBox2.Text + "',otchestvo='" + textBox3.Text + "',data_oformlenya='"
                + textBox4.Text + "',elektronnaya_pochta='" + textBox5.Text + "',nomer_telefona='" + textBox6.Text + "',rachetnyu_chet='" + textBox7.Text + "',nazvanie_banka='" + textBox8.Text + "',kod_sotrudnika='" + textBox9.Text + "',kod_centra='" + comboBox1.Text + "',kod_pola='" + comboBox2.Text + "' where kod_klienta = '" + textBox11.Text + "'";
                cmd.ExecuteNonQuery();
                con.Close();
                GetList();
                MessageBox.Show(string.Format("Успешное изменение данных в таблице"), "Сообщение");
            }
            else
            {
                MessageBox.Show("Текстовое поле 'Код клиента' не заполнено", "Ошибка");
            }
        }
        // Функция "Добавить"
        private void button2_Click(object sender, RoutedEventArgs e)
        {
            if ((textBox1.Text != "") && (textBox2.Text != "") && (textBox3.Text != "") && (textBox4.Text != "") && (textBox5.Text != "") && (textBox6.Text != "") && (textBox7.Text != "") && (textBox8.Text != "") && (textBox9.Text != "") && (comboBox1.Text != "") && (comboBox2.Text != ""))
            {
                cmd = new SqlCommand();
                cmd.Connection = con;
                con.Open();
                cmd.CommandText = "INSERT INTO Diplom_Klient (familiya,imya,otchestvo,data_oformlenya,elektronnaya_pochta,nomer_telefona,rachetnyu_chet,nazvanie_banka,kod_sotrudnika,kod_centra,kod_pola) values (" + textBox1.Text + ",'" + textBox2.Text + "','" + textBox3.Text + "','" + textBox4.Text + "','" + textBox5.Text + "','" + textBox6.Text + "','" + textBox7.Text + "','" + textBox8.Text + "','" + textBox9.Text + "','" + comboBox1.Text + "','" + comboBox2.Text + "')";
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
        // Функция "Вывод в Excel"
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
        // Функция "Поиск по фамилии"
        private void button6_Click(object sender, RoutedEventArgs e)
        {
            if (textBox15.Text != "")
            {
                string filter = textBox15.Text;
                DataView view = dataGridView1.ItemsSource as DataView;
                if (view != null)
                {
                    view.RowFilter = $"familiya LIKE '%{filter}%'";
                }
            }
            else
            {
                MessageBox.Show("Текстовое поле 'Поиск по фамилии' не заполнено", "Ошибка");
            }
        }
        // Функция "Очистить"
        private void button7_Click(object sender, RoutedEventArgs e)
        {
            textBox15.Clear();
            adm_klient form = new adm_klient();
            form.Show();
            this.Close();
        }
        // Функция "Переход на форму 'БД Приемы'"
        private void button4_Click(object sender, RoutedEventArgs e)
        {
            adm_priem form = new adm_priem();
            form.Show();
        }
        // Функция "Выйти"
        private void button1_Click(object sender, RoutedEventArgs e)
        {
            glavnaya_administrator form = new glavnaya_administrator();
            form.Show();
            this.Close();
        }
        // Функция "Сортировка по возрастанию"
        private void button8_Click(object sender, RoutedEventArgs e)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                SqlCommand command = new SqlCommand("SELECT * FROM Diplom_Klient ORDER BY data_oformlenya ASC", connection);
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
                SqlCommand command = new SqlCommand("SELECT * FROM Diplom_Klient ORDER BY data_oformlenya DESC", connection);
                SqlDataReader reader = command.ExecuteReader();
                System.Data.DataTable table = new System.Data.DataTable();
                table.Load(reader);
                dataGridView1.ItemsSource = table.DefaultView;
                reader.Close();
            }
        }
        // Функция "Перезагрузка"
        private void button10_Click(object sender, RoutedEventArgs e)
        {
            adm_klient form = new adm_klient();
            form.Show();
            this.Close();
        }
    }
}
