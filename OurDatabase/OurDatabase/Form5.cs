using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace OurDatabase
{
    public partial class Form5 : Form
    {
        public static string connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb;"; //строка подключения к базе данных
        private OleDbConnection myConnection; //подключение
        private OleDbDataAdapter dataadapter; //адаптер для таблицы
        List<int> codes_of_country = new List<int>(); //коды стран, городов, континентов и певцов. По ним идет обращение к базе
        List<int> codes_of_city = new List<int>();
        List<int> codes_of_singer = new List<int>();
        List<int> codes_of_continents = new List<int>();

        int nContinent = 0; // номер выбранного города, страны, континента, певца. Формирует запрос where
        int nCountry = 0;
        int nCity = 0;
        int nGroup = 0;
        string sw = ""; //нач строка
        public Form5()
        {
            InitializeComponent();
        }

       


        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();

        }

        private string Function(int nGroup, int nCity, int nCountry, int nContinent)
        {
            //функция, формирующая запрос WHERE для установки фильтра
            
            if (nGroup == 0)
            {
                sw = " WHERE [SingerGroup].[Code_Singer] >0 "; //для правильного создания запроса, если певец не выбран
            }
            else
            {
                sw = " WHERE [SingerGroup].[Code_Singer] = " + nGroup.ToString(); // вывод тех полей, где номер певца = номеру выбранного 
            }

            if (nCity > 0) //если выбран еще город
            {
                sw += " AND [City].[Code_City] = " + nCity.ToString();  // те поля, где город=выбранный город
            }
            if (nCountry > 0)
            {
                sw += " AND [Country].[Code_Country] = " + nCountry.ToString();
            }
            if (nContinent > 0)
            {
                sw += " AND [Continent].[Code_Continent] = " + nContinent.ToString();
            }
            return sw;
        }

        private void Count_title(List<int> codes_of_singer, int count)
        {
            int c = 0;
            int cc = 0;
            for (int i = 0; i < count; i++)
            {
                string count_songs = "SELECT COUNT([Title]) FROM [Single] WHERE [Single].[NumSingerGroup] = " + codes_of_singer[i].ToString();
                OleDbCommand songs = new OleDbCommand(count_songs, myConnection);
                OleDbDataReader read_songs = songs.ExecuteReader();
                while (read_songs.Read())
                {
                    c = Convert.ToInt32(read_songs[0]);

                }
                read_songs.Close();
                cc += c;
            }

            label5.Text = cc.ToString();
        }


        private void button4_Click(object sender, EventArgs e)
        {
           
            string connectData = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb;"; //подключение
            string sql = "SELECT [Title] as Название,[ArtistName] as Группа,[CityName] as Город, " +
                "[CountryName] as Страна,[ContinentName] as Континент"; //часть главного запроса для вывода полей

            string tables = " FROM [Single],[SingerGroup],[City],[Country],[Continent] "; //с каких таблиц берем данные

            string sw = Function(nGroup, nCity, nCountry, nContinent); //получаем условие WHERE из фильтров
            string origin = " AND [SingerGroup].[Code_Singer] = [Single].[NumSingerGroup] AND [City].[Code_City] = [SingerGroup].[NumCity]" +
                " AND [Country].[Code_Country] = [City].[NumCountry]" +
                " AND [Country].[NumContinent] = [Continent].[Code_Continent]"; //соединение таблиц для лииквидации дублирования
            sw += origin;
            sw += " ORDER BY [Title]"; //сортировка

            sql += tables + sw;
            myConnection = new OleDbConnection(connectData);

            dataadapter = new OleDbDataAdapter(sql, myConnection);

            myConnection.Open(); //создание таблицы 
            DataSet ds = new DataSet();
            dataadapter.Fill(ds);
            table.DataSource = ds.Tables[0];
            table.ReadOnly = true;

        }

        private void Form5_Load(object sender, EventArgs e)
        {
            //запрос соединения таблиц 
            string connectData = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb;";
            string sql = "SELECT [Title] as Название,[ArtistName] as Группа,[CityName] as Город,[CountryName] as Страна,[ContinentName] as Континент" +
               " FROM [Single],[SingerGroup],[City],[Country],[Continent]" +
               " WHERE [SingerGroup].[Code_Singer] = [Single].[NumSingerGroup] AND [City].[Code_City] = [SingerGroup].[NumCity] " +
               "AND [Country].[Code_Country] = [City].[NumCountry]" +
               " AND [Country].[NumContinent] = [Continent].[Code_Continent]";

            myConnection = new OleDbConnection(connectData);
            dataadapter = new OleDbDataAdapter(sql, myConnection);
            myConnection.Open();
            DataSet ds = new DataSet();
            dataadapter.Fill(ds);
            table.DataSource = ds.Tables[0];
            table.ReadOnly = true;


            Virtual_Table(); //начальная инициализация комбо_боксов


            label6.Text = MySinger.Items.Count.ToString();
            label7.Text = MyCity.Items.Count.ToString();
            label8.Text = comboBox3.Items.Count.ToString();
            label9.Text = MyContinent.Items.Count.ToString();
            button4.Enabled = true;
        }

        private void Virtual_Table()
        {
            //виртульная таблица для регулирования комбо_боксов
            string sql2 = "SELECT [Continent].[ContinentName], [Country].[CountryName],[City].[CityName],[SingerGroup].[ArtistName]," +
                "[Continent].[Code_Continent],[Country].[Code_Country],[City].[Code_City],[SingerGroup].[Code_Singer] " +
                "FROM [SingerGroup] INNER JOIN " +
                "([City] INNER JOIN " +
                "([Country] INNER JOIN [Continent] ON [Continent].[Code_Continent] = [Country].[NumContinent]) ON [Country].[Code_Country] = [City].[NumCountry]) " +
                "ON [City].[Code_City] = [SingerGroup].[NumCity] ";
            sql2 += Function(nGroup, nCity, nCountry, nContinent); //получаем условие WHERE из фильтров
            OleDbCommand combo = new OleDbCommand(sql2, myConnection);
            OleDbDataReader que_reader_c = combo.ExecuteReader();
            while (que_reader_c.Read())
            {

                MyContinent.Items.Add(que_reader_c[0].ToString());
                comboBox3.Items.Add(que_reader_c[1].ToString());
                MyCity.Items.Add(que_reader_c[2].ToString());
                MySinger.Items.Add(que_reader_c[3].ToString());
                codes_of_continents.Add(Convert.ToInt32(que_reader_c[4]));
                codes_of_country.Add(Convert.ToInt32(que_reader_c[5]));
                codes_of_city.Add(Convert.ToInt32(que_reader_c[6]));
                codes_of_singer.Add(Convert.ToInt32(que_reader_c[7]));
            }
            que_reader_c.Close();

            //убрать дубли в комбо боксах
            object[] items1 = MyContinent.Items.OfType<String>().Distinct().ToArray();
            MyContinent.Items.Clear();
            MyContinent.Items.AddRange(items1);

            object[] items2 = comboBox3.Items.OfType<String>().Distinct().ToArray();
            comboBox3.Items.Clear();
            comboBox3.Items.AddRange(items2);

            object[] items3 = MyCity.Items.OfType<String>().Distinct().ToArray();
            MyCity.Items.Clear();
            MyCity.Items.AddRange(items3);

            // убрать дублирующиес коды
            int[] items11 = codes_of_continents.OfType<int>().Distinct().ToArray();
            codes_of_continents.Clear();
            codes_of_continents.AddRange(items11);

            int[] items12 = codes_of_country.OfType<int>().Distinct().ToArray();
            codes_of_country.Clear();
            codes_of_country.AddRange(items12);

            int[] items13 = codes_of_city.OfType<int>().Distinct().ToArray();
            codes_of_city.Clear();
            codes_of_city.AddRange(items13);
        }
        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }
        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            nCountry = codes_of_country[comboBox3.SelectedIndex];
            Virtual_Table();
           
        }

        private void MyContinent_SelectedIndexChanged(object sender, EventArgs e)
        {
            nContinent = codes_of_continents[MyContinent.SelectedIndex];
           
           

        }

        private void MyCity_SelectedIndexChanged(object sender, EventArgs e)
        {
            nCity = codes_of_city[MyCity.SelectedIndex];
        }

        private void table_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }


        private void MySinger_Click(object sender, EventArgs e)
        {

        }

        private void MySinger_SelectedIndexChanged_1(object sender, EventArgs e)
        {

           
            nGroup = codes_of_singer[MySinger.SelectedIndex];


        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form5_Load(this, null);
            codes_of_continents.Clear();
            codes_of_city.Clear();
            codes_of_singer.Clear();
            codes_of_country.Clear();
            MySinger.Items.Clear();
            MyContinent.Items.Clear();
            comboBox3.Items.Clear();
            MyCity.Items.Clear();

            MySinger.Text = "";
            MyContinent.Text = "";
            comboBox3.Text = "";
            MyCity.Text = "";

           

            nContinent = 0;
            nCountry = 0;
            nCity = 0;
            nGroup = 0;
           


            myConnection = new OleDbConnection(connectString);
            myConnection.Open();

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExportToExcel();
        }

        private void PopulateRows()
        {
            for (int i = 1; i <= 10; i++)
            {
                DataGridViewRow row =
                    (DataGridViewRow)table.RowTemplate.Clone();

                row.CreateCells(table, string.Format("Song{0}", i),
                    string.Format("Singer{0}", i), string.Format("City{0}", i), string.Format("Country{0}", i), string.Format("Continent{0}", i));

                table.Rows.Add(row);

            }
        }

        private void ExportToExcel()
        {

            Microsoft.Office.Interop.Excel._Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = excel.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            try
            {
                worksheet = workbook.ActiveSheet;
                worksheet.Name = "Songs";

                int cellRowIndex = 1;
                int cellColumnIndex = 1;


                for (int i = 0; i < table.Rows.Count - 1; i++)
                {
                    for (int j = 0; j <table.Columns.Count; j++)
                    {
                        // Excel index starts from 1,1. As first Row would have the Column headers, adding a condition check.
                        if (cellRowIndex == 1)
                        {
                            worksheet.Cells[cellRowIndex, cellColumnIndex] = table.Columns[j].HeaderText;
                        }
                        else
                        {
                            worksheet.Cells[cellRowIndex, cellColumnIndex] = table.Rows[i].Cells[j].Value.ToString();
                        }
                        cellColumnIndex++;
                    }
                    cellColumnIndex = 1;
                    cellRowIndex++;
                }

                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                saveDialog.FilterIndex = 2;

                if (saveDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    workbook.SaveAs(saveDialog.FileName);
                    MessageBox.Show("Export Successful");
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                excel.Quit();
                workbook = null;
                excel = null;
            }
        }
    }
}
