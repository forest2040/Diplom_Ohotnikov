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
    /// Логика взаимодействия для sotrudniki.xaml
    /// </summary>
    public partial class sotrudniki
    {
        private string connectionString = "data source = stud-mssql.sttec.yar.ru,38325; initial catalog = user57_db; user id = user57_db; password = user57";
        SqlCommand cmd;
        SqlConnection con;
        SqlDataAdapter da;
        DataSet ds;
        public sotrudniki()
        {
            InitializeComponent();
            GetList();
            LoadData();
        }
        void GetList()
        {
            con = new SqlConnection(@"data source = stud-mssql.sttec.yar.ru,38325; initial catalog = user57_db; user id = user57_db; password = user57");
            con.Open();
            da = new SqlDataAdapter("SELECT * FROM Diplom_Sotrudnik", con);
            ds = new DataSet();
            da.Fill(ds, "Diplom_Sotrudnik");
            DataView dview = new DataView(ds.Tables["Diplom_Sotrudnik"]);
            dataGridView1.ItemsSource = dview;
            con.Close();
        }
        private void LoadData()
        {
            string query = "SELECT * FROM Diplom_Sotrudnik";
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
            if ((textBox1.Text != "") && (textBox2.Text != "") && (textBox3.Text != "") && (textBox4.Text != "") && (comboBox3.Text != "") && (textBox6.Text != "") && (textBox7.Text != "") && (textBox8.Text != "") && (textBox9.Text != "") && (textBox10.Text != "") && (textBox11.Text != "") && (textBox12.Text != "") && (comboBox1.Text != "") && (comboBox2.Text != ""))
            {
                cmd = new SqlCommand();
                cmd.Connection = con;
                con.Open();
                cmd.CommandText = "INSERT INTO Diplom_Sotrudnik (familiya,imya,otchestvo,den_rozhdenya,dolzhnost,specializacya,kfalifikacya,elektronnaya_pochta,nomer_telefona,adres,login,parol,kod_centra, kod_pola) values (" + textBox1.Text + ",'" + textBox2.Text + "','" + textBox3.Text + "','" + textBox4.Text + "','" + comboBox3.Text + "','" + textBox6.Text + "','" + textBox7.Text + "','" + textBox8.Text + "','" + textBox9.Text + "','" + textBox10.Text + "','" + textBox11.Text + "','" + textBox12.Text + "','" + comboBox1.Text + "','" + comboBox2.Text + "')";
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
            if (textBox14.Text != "")
            {
                cmd = new SqlCommand();
                cmd.Connection = con;
                con.Open();
                cmd.CommandText = "UPDATE Diplom_Sotrudnik SET familiya = '" + textBox1.Text + "',imya='" + textBox2.Text + "',otchestvo='" + textBox3.Text +  "',den_rozhdenya='"
                + textBox4.Text + "',dolzhnost='" + comboBox3.Text + "',specializacya='" + textBox6.Text + "',kfalifikacya='" + textBox7.Text + "',elektronnaya_pochta='" + textBox8.Text + "',nomer_telefona='" + textBox9.Text + "',adres='" + textBox10.Text + "',login='" + textBox11.Text + "',parol='" + textBox12.Text + "',kod_centra='" + comboBox1.Text + "',kod_pola='" + comboBox2.Text + "' where kod_sotrudnika = '" + textBox14.Text + "'";
                cmd.ExecuteNonQuery();
                con.Close();
                GetList();
                MessageBox.Show(string.Format("Успешное изменение данных в таблице"), "Сообщение");
            }
            else
            {
                MessageBox.Show("Текстовое поле 'Код сотрудника' не заполнено", "Ошибка");
            }
        }
        // Функция "Удалить"
        private void button4_Click(object sender, RoutedEventArgs e)
        {
            if (textBox14.Text != "")
            {
                cmd = new SqlCommand();
                cmd.Connection = con;
                con.Open();
                cmd.CommandText = "DELETE FROM Diplom_Sotrudnik WHERE kod_sotrudnika = '" + textBox14.Text + "'";
                cmd.ExecuteNonQuery();
                con.Close();
                GetList();
                MessageBox.Show(string.Format("Успешное удаление данных в таблице"), "Сообщение");
                textBox1.Clear();
                textBox2.Clear();
                textBox3.Clear();
                textBox4.Clear();
                textBox6.Clear();
                textBox7.Clear();
                textBox8.Clear();
                textBox9.Clear();
                textBox10.Clear();
                textBox11.Clear();
                textBox12.Clear();
                textBox14.Clear();
                comboBox1.Items.Clear();
                comboBox2.Items.Clear();
                comboBox3.Items.Clear();
            }
            else
            {
                MessageBox.Show("Текстовое поле 'Код сотрудника' не заполнено", "Ошибка");
            }
        }
        // Функция "Выйти"
        private void button1_Click(object sender, RoutedEventArgs e)
        {
            glavnaya_director form = new glavnaya_director();
            form.Show();
            this.Close();
        }
        // Функция "Переход на форму 'БД Статус'"
        private void button5_Click(object sender, RoutedEventArgs e)
        {
            status form = new status();
            form.Show();
        }
        // Вывод в Excel
        private void button6_Click(object sender, RoutedEventArgs e)
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
        // Функция "Перезагрузка формы"
        private void button7_Click(object sender, RoutedEventArgs e)
        {
            sotrudniki form = new sotrudniki();
            form.Show();
            this.Close();
        }
        // Функция "Поиск по фамилии"
        private void button8_Click(object sender, RoutedEventArgs e)
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
        private void button9_Click(object sender, RoutedEventArgs e)
        {
            textBox15.Clear();
            sotrudniki form = new sotrudniki();
            form.Show();
            this.Close();
        }
        // Функция "Сортировка по возрастанию"
        private void button10_Click(object sender, RoutedEventArgs e)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                SqlCommand command = new SqlCommand("SELECT * FROM Diplom_Sotrudnik ORDER BY den_rozhdenya ASC", connection);
                SqlDataReader reader = command.ExecuteReader();
                System.Data.DataTable table = new System.Data.DataTable();
                table.Load(reader);
                dataGridView1.ItemsSource = table.DefaultView;
                reader.Close();
            }
        }
        // Функция "Сортировка по убыванию"
        private void button11_Click(object sender, RoutedEventArgs e)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                SqlCommand command = new SqlCommand("SELECT * FROM Diplom_Sotrudnik ORDER BY den_rozhdenya DESC", connection);
                SqlDataReader reader = command.ExecuteReader();
                System.Data.DataTable table = new System.Data.DataTable();
                table.Load(reader);
                dataGridView1.ItemsSource = table.DefaultView;
                reader.Close();
            }
        }
    }
    }
