using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Program
{
    public partial class Нагрузка : Form
    {
        public Нагрузка()
        {
            InitializeComponent();
        }

        OleDbConnection con;    //Строка соединения с БД
        OleDbCommand SqlCom;    //Переменная для Sql-запросов
        OleDbCommand SqlComV;   //Врачи
        OleDbCommand SqlComP;   //Пациенты
        DataTable DT;           //Таблица для хранения результатов запроса
        OleDbDataAdapter DA;    // Адаптер для заполнения таблицы после запроса
        bool ifcon = false;     //Флаг срединения с базой данных  

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void ShowList()
        {
            //Вывод списка в таблицу DataGridView  
            DT = new DataTable();  //Создаем заново таблицу       
            SqlCom.ExecuteNonQuery(); //Выполняем запрос
            DA = new OleDbDataAdapter(SqlCom); //Через адаптер получаем результаты запроса
            DA.Fill(DT); // Заполняем таблицу результами
            dataGridView1.DataSource = DT; // Привязываем DataGridView1 к источнику данных
            dataGridView1.Columns[0].Visible = false;    //Код преподавателя невидимый
            dataGridView1.Columns[1].Visible = false;  //Код группы невидимый
        }

        private void ShowVP() 
        {
            SqlComV = new OleDbCommand("SELECT * FROM Нагрузка", con);
            OleDbDataReader dataReaderV = SqlComV.ExecuteReader();     //Создать объект для чтения и выполнить команду
                                                                       //Таблица прочитана в dataReaderV (виртуальная таблица)
            comboBox1.Items.Clear();    //Очистка списка ComboBox1 (ФИО врачей)

            while (dataReaderV.Read())  //Пока не конец виртуальной таблицы
            {
                comboBox1.Items.Add(dataReaderV.GetValue(2)); //Номер группы
                comboBox2.Items.Add(dataReaderV.GetValue(0)); // id преподавателя
                comboBox3.Items.Add(dataReaderV.GetValue(1)); // id группы
            }
            dataReaderV.Close();    //Закрыть объект чтения
            comboBox1.SelectedIndex = -1;
        }
        private void Нагрузка_Load(object sender, EventArgs e)
        {
            try
            {
                con = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=Распределение учебной нагрузки.accdb");
                con.Open();     //Открыть базу данных
                ifcon = true;   //Флаг поднят. Соединение с базой данных прошло успешно.          
                ShowVP();       //Вызов процедуры формирования списков врачей и пациентов
                                // Указываем строку запроса и привязываем к соединению
                SqlCom = new OleDbCommand("SELECT * FROM Нагрузка", con);
                ShowList(); //Вызов процедуры пручения списка обращений

            }
            catch (System.Exception err)
            {
                MessageBox.Show(err.Message, "ОШИБКА ДОСТУПА К БАЗЕ ДАННЫХ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Нагрузка_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (ifcon) con.Close(); //Закрыть базу данных, если она была успешно открыта
        }
        private void button1_Click(object sender, EventArgs e)
        {
            
            DateTime dat = dateTimePicker1.Value;
            DateTime dat2 = dateTimePicker2.Value;
            int idp = Convert.ToInt32(comboBox2.SelectedItem);  //Получить код преподавателя
            int idg = Convert.ToInt32(comboBox3.SelectedItem);  //Получить код группы
            //Приведение даты к формату dd.mm.yyyy
            int d = dat.Day;
            int m = dat.Month;
            int y = dat.Year;
            int d2 = dat2.Day;
            int m2 = dat2.Month;
            int y2 = dat2.Year;
            string dat1 = "";
            string dats2 = "";
            if (m < 10)
                dat1 = Convert.ToString(d) + ".0" + Convert.ToString(m) + "." + Convert.ToString(y);
            else
                dat1 = Convert.ToString(d) + "." + Convert.ToString(m) + "." + Convert.ToString(y);
            if (m2 < 10)
                dats2 = Convert.ToString(d2) + ".0" + Convert.ToString(m2) + "." + Convert.ToString(y2);
            else
                dats2 = Convert.ToString(d2) + "." + Convert.ToString(m2) + "." + Convert.ToString(y2);
            OleDbCommand SqlCom1 = new OleDbCommand();
            SqlCom1.CommandText = "INSERT INTO [Нагрузка] ([Код преподавателя], [Код предмета], [Номер группы], Дата_начала, Дата_окончания) VALUES (@idp, @idg, @group, @d1, @d2)";
            SqlCom1.Parameters.Clear();
            SqlCom1.Parameters.AddWithValue("@idp", idp);
            SqlCom1.Parameters.AddWithValue("@idg", idg);
            SqlCom1.Parameters.AddWithValue("@d1", dat1);
            SqlCom1.Parameters.AddWithValue("@d2", dats2);
            SqlCom1.Parameters.AddWithValue("@group", comboBox1.Text);
            SqlCom1.Connection = con;
            SqlCom1.ExecuteScalar(); //Выполняем запрос
            ShowList();
            ShowVP();
            //ClearAll();
            MessageBox.Show("Запись добавлена.", "ДОБАВЛЕНИЕ ЗАПИСИ", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text.Trim() == "")
                MessageBox.Show("Выберите строку в таблице или нажмите ДОБАВИТЬ.",
                    "ОШИБКА В ОПЕРАЦИИ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                DateTime dat = dateTimePicker1.Value;
                DateTime dat2 = dateTimePicker2.Value;
                int idp = Convert.ToInt32(comboBox2.SelectedItem);  //Получить код преподавателя
                int idg = Convert.ToInt32(comboBox3.SelectedItem);  //Получить код группы
                                                                    //Приведение даты к формату dd.mm.yyyy
                int d = dat.Day;
                int m = dat.Month;
                int y = dat.Year;
                int d2 = dat2.Day;
                int m2 = dat2.Month;
                int y2 = dat2.Year;
                string dat1 = "";
                string dats2 = "";
                if (m < 10)
                    dat1 = Convert.ToString(d) + ".0" + Convert.ToString(m) + "." + Convert.ToString(y);
                else
                    dat1 = Convert.ToString(d) + "." + Convert.ToString(m) + "." + Convert.ToString(y);
                if (m2 < 10)
                    dats2 = Convert.ToString(d2) + ".0" + Convert.ToString(m2) + "." + Convert.ToString(y2);
                else
                    dats2 = Convert.ToString(d2) + "." + Convert.ToString(m2) + "." + Convert.ToString(y2);
                OleDbCommand SqlCom1 = new OleDbCommand();
                SqlCom1.CommandText = "UPDATE [Нагрузка] SET [Код преподавателя] = @idp, [Код предмета] = @idg, [Номер группы] = @group, Дата_начала = @d1, Дата_окончания = d2 WHERE [Код предмета] = @idp ";
                SqlCom1.Parameters.Clear();
                SqlCom1.Parameters.AddWithValue("@idp", idp);
                SqlCom1.Parameters.AddWithValue("@idg", idg);
                SqlCom1.Parameters.AddWithValue("@d1", dat1);
                SqlCom1.Parameters.AddWithValue("@d2", dats2);
                SqlCom1.Parameters.AddWithValue("@group", comboBox1.Text);
                SqlCom1.Connection = con;
                SqlCom1.ExecuteScalar(); //Выполняем запрос
                ShowList();
                ShowVP();
                //ClearAll();
                MessageBox.Show("Запись изменена.", "ИЗМЕНЕНИЕ ЗАПИСИ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text.Trim() == "")
            {
                MessageBox.Show("Выберите строку для удаления!",
                    "ОШИБКА В ОПЕРАЦИИ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                OleDbCommand SqlCom1 = new OleDbCommand();
                SqlCom1.CommandText = "DELETE FROM [Нагрузка] WHERE [Код предмета] = @idg";
                SqlCom1.Parameters.Clear();
                SqlCom1.Parameters.AddWithValue("@idg", comboBox1.Text);
                SqlCom1.Connection = con;
                SqlCom1.ExecuteScalar(); //Выполняем запрос
                ShowList();
                MessageBox.Show("Запись удалена.", "УДАЛЕНИЕ ЗАПИСИ",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
    
        }
    }

}
