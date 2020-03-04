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
using System.Threading;

namespace OurDatabase
{
    public partial class Form5 : Form
    {
        public static string connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb;"; 
        private OleDbConnection myConnection; 
        private OleDbDataAdapter dataadapter; 
        List<int> codes_of_country = new List<int>(); 
        List<int> codes_of_city = new List<int>();
        List<int> codes_of_singer = new List<int>();
        List<int> codes_of_continents = new List<int>();

        Mutex mm = new Mutex();
        
        string singer_table = "[SingerGroup].[ArtistName],[SingerGroup].[Code_Singer]";
        string city_table = "[City].[CityName],[City].[Code_City]";
        string country_table = "[Country].[CountryName],[Country].[Code_Country]";
        string continent_table = "[Continent].[ContinentName],[Continent].[Code_Continent]";

        string table_query = " FROM [SingerGroup] INNER JOIN " +
              "([City] INNER JOIN " +
              "([Country] INNER JOIN [Continent] ON [Continent].[Code_Continent] = [Country].[NumContinent]) ON [Country].[Code_Country] = [City].[NumCountry]) " +
              "ON [City].[Code_City] = [SingerGroup].[NumCity] ";
        
        int nContinent = 0; 
        int nCountry = 0;
        int nCity = 0;
        int nGroup = 0;

        string cc = "";
        
        int const_list = 0;

        string sw = ""; 
    
        public Form5()
        {
            InitializeComponent();
        }
        private string Regule_One_List(int k)
        {
           
            string uniq_field = "";
            if(k==1) 
            {
                uniq_field += singer_table + "," + city_table + "," + country_table;
            }
            else if(k==2) 
            {
                uniq_field += singer_table + "," + city_table + "," + continent_table;
            }
            else if(k==3) 
            {
                uniq_field += singer_table + "," + country_table + "," + continent_table;
            }
            else if(k==4) 
            {
                uniq_field += city_table + "," + country_table + "," + continent_table;
            }
            return uniq_field;
        }

        
       


        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();

        }




        private string Get_Query(int a, int b, int c, int d,int uniq_one)
        {
            string field = Regule_One_List(uniq_one); 
            string sql2 = "SELECT DISTINCT " + field; 
            sql2 += table_query;
            string where = Function(a, b, c, d); 
            sql2 += where;     
            return sql2;
        }



        private string Function(int nGroup,int nCity,int nCountry,int nContinent)
        {
            //функция, формирующая запрос WHERE для установки фильтра
            
            if (nGroup == 0)
            {
                sw = " WHERE [SingerGroup].[Code_Singer] > 0 "; //для правильного создания запроса, если певец не выбран
            }
            else
            {
                sw = " WHERE [SingerGroup].[Code_Singer] = " + nGroup.ToString(); // вывод тех полей, где номер певца = номеру выбранного 
            }
            if (nContinent > 0)
            {
                sw += " AND [Continent].[Code_Continent] = " + nContinent.ToString();
            }
            
            if (nCountry > 0)
            {
                sw += " AND [Country].[Code_Country] = " + nCountry.ToString();
            }
            
            if (nCity > 0) //если выбран еще город
            {
                sw += " AND [City].[Code_City] = " + nCity.ToString();  // те поля, где город=выбранный город
            }
            
            return sw;
        }
        private void button4_Click(object sender, EventArgs e)
        {
           
            string connectData = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb;"; 
            string sql = "SELECT [Title] as Название,[ArtistName] as Группа,[CityName] as Город, " +
                "[CountryName] as Страна,[ContinentName] as Континент"; 

            string tables = " FROM [Single],[SingerGroup],[City],[Country],[Continent] "; 

            string sw = Function(nGroup, nCity, nCountry,nContinent); 
            string origin = " AND [SingerGroup].[Code_Singer] = [Single].[NumSingerGroup] AND [City].[Code_City] = [SingerGroup].[NumCity]" +
                " AND [Country].[Code_Country] = [City].[NumCountry]" +
                " AND [Country].[NumContinent] = [Continent].[Code_Continent]"; 
            sw += origin;
            sw += " ORDER BY [Title]"; 

            sql += tables + sw;
            myConnection = new OleDbConnection(connectData);

            dataadapter = new OleDbDataAdapter(sql, myConnection);

            myConnection.Open();  
            DataSet ds = new DataSet();
            dataadapter.Fill(ds);
            table.DataSource = ds.Tables[0];


            table.ReadOnly = true;

            int count = table.RowCount-1;
            label5.Text = count.ToString();
            if (nGroup > 0)
            {
                label6.Text = 1.ToString();
            }
            else
            {
                label6.Text = MySinger.Items.Count.ToString();
            }
            
            if (nCity > 0)
            {
                label7.Text = 1.ToString();
            }
            else
            {
               
                label7.Text = MyCity.Items.Count.ToString();

            }
            if(nCountry > 0)
            {
                label8.Text = 1.ToString();
            }
            else
            {
                label8.Text = comboBox3.Items.Count.ToString();
                
            }
            if (nContinent > 0)
            {
                label9.Text = 1.ToString();
            }
            else
            {
                label9.Text = MyContinent.Items.Count.ToString();
            }

           


        }
        
        private void Form5_Load(object sender, EventArgs e)
        {
            
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


            string sql2 = "SELECT DISTINCT " + singer_table + "," + city_table + "," + country_table + "," + continent_table;
            sql2 += table_query;
            OleDbCommand combo = new OleDbCommand(sql2, myConnection);
            OleDbDataReader que_reader_c = combo.ExecuteReader();
            while (que_reader_c.Read())
            {

                MyContinent.Items.Add(que_reader_c[6].ToString());
                comboBox3.Items.Add(que_reader_c[4].ToString());
                MyCity.Items.Add(que_reader_c[2].ToString());
                MySinger.Items.Add(que_reader_c[0].ToString());
                codes_of_continents.Add(Convert.ToInt32(que_reader_c[7]));
                codes_of_country.Add(Convert.ToInt32(que_reader_c[5]));
                codes_of_city.Add(Convert.ToInt32(que_reader_c[3]));
                codes_of_singer.Add(Convert.ToInt32(que_reader_c[1]));
            }
            que_reader_c.Close();
            Clear_Info();
            int count = table.RowCount - 1;
            label5.Text = count.ToString();
            label6.Text = MySinger.Items.Count.ToString();
            label7.Text = MyCity.Items.Count.ToString();
            label8.Text = comboBox3.Items.Count.ToString();
            label9.Text = MyContinent.Items.Count.ToString();
            button4.Enabled = true;
        }

       

        private void Clear_Info()
        {
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
            const_list = 2;
            string query = Get_Query(nGroup, nCity, nCountry, nContinent, const_list);
            Changing_Combo_Box(query, const_list);

           
           

           
        }

        private void MyContinent_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            
            nContinent = codes_of_continents[MyContinent.SelectedIndex];
            const_list = 1;
            string query = Get_Query(nGroup, nCity, nCountry, nContinent, const_list);
            Changing_Combo_Box(query, const_list);
           
        }

        private void MyCity_SelectedIndexChanged(object sender, EventArgs e)
        {

            
            nCity = codes_of_city[MyCity.SelectedIndex];
            const_list = 3;
            string query = Get_Query(nGroup, nCity, nCountry, nContinent, const_list);
            Changing_Combo_Box(query, const_list);
            

           

        }

        private void table_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        private void Changing_Combo_Box(string query_for_change,int this_list)
        {
            if(this_list==1) 
            {
                
                OleDbCommand combo = new OleDbCommand(query_for_change, myConnection);
                OleDbDataReader que_reader_c = combo.ExecuteReader();

                List<string> item = new List<string>();
                List<int> h_item = new List<int>();

                List<string> item1 = new List<string>();
                List<int> h_item1= new List<int>();

                List<string> item2 = new List<string>();
                List<int> h_item2 = new List<int>();

                MySinger.Items.Clear();
                MyCity.Items.Clear();
                comboBox3.Items.Clear();

                codes_of_singer.Clear();
                codes_of_city.Clear();
                codes_of_country.Clear();
                while (que_reader_c.Read())
                {
                    item.Add(que_reader_c[4].ToString());
                    h_item.Add(Convert.ToInt32(que_reader_c[5]));
                    item1.Add(que_reader_c[2].ToString());
                    h_item1.Add(Convert.ToInt32(que_reader_c[3]));
                    item2.Add(que_reader_c[0].ToString());
                    h_item2.Add(Convert.ToInt32(que_reader_c[1]));

                   
                }
                que_reader_c.Close();

                object[] country = item.OfType<string>().Distinct().ToArray();
                comboBox3.Items.AddRange(country);

                int[] country_code = h_item.OfType<int>().Distinct().ToArray();
                codes_of_country.AddRange(country_code);

                object[] city = item1.OfType<string>().Distinct().ToArray();
                MyCity.Items.AddRange(city);

                int[] city_code = h_item1.OfType<int>().Distinct().ToArray();
                codes_of_city.AddRange(city_code);

                object[] singer = item2.OfType<string>().Distinct().ToArray();
                MySinger.Items.AddRange(singer);

                int[] singer_code = h_item2.OfType<int>().Distinct().ToArray();
                codes_of_singer.AddRange(singer_code);

               

            }
            if (this_list == 2) 
            {
                
                OleDbCommand combo = new OleDbCommand(query_for_change, myConnection);
                OleDbDataReader que_reader_c = combo.ExecuteReader();

                List<string> item = new List<string>();
                List<int> h_item = new List<int>();

                List<string> item1 = new List<string>();
                List<int> h_item1 = new List<int>();

                List<string> item2 = new List<string>();
                List<int> h_item2 = new List<int>();



                MySinger.Items.Clear();
                MyCity.Items.Clear();
                MyContinent.Items.Clear();

                codes_of_singer.Clear();
                codes_of_city.Clear();
                codes_of_continents.Clear();
                while (que_reader_c.Read())
                {
                    item.Add(que_reader_c[4].ToString());
                    h_item.Add(Convert.ToInt32(que_reader_c[5]));
                    item1.Add(que_reader_c[2].ToString());
                    h_item1.Add(Convert.ToInt32(que_reader_c[3]));
                    item2.Add(que_reader_c[0].ToString());
                    h_item2.Add(Convert.ToInt32(que_reader_c[1]));


                   
                }
                que_reader_c.Close();

                object[] continents = item.OfType<string>().Distinct().ToArray();
                MyContinent.Items.AddRange(continents);

                int[] cont_code = h_item.OfType<int>().Distinct().ToArray();
                codes_of_continents.AddRange(cont_code);

                object[] city = item1.OfType<string>().Distinct().ToArray();
                MyCity.Items.AddRange(city);

                int[] city_code = h_item1.OfType<int>().Distinct().ToArray();
                codes_of_city.AddRange(city_code);

                object[] singer = item2.OfType<string>().Distinct().ToArray();
                MySinger.Items.AddRange(singer);

                int[] singer_code = h_item2.OfType<int>().Distinct().ToArray();
                codes_of_singer.AddRange(singer_code);

                

            }
            if (this_list == 3) 
            {
                
                OleDbCommand combo = new OleDbCommand(query_for_change, myConnection);
                OleDbDataReader que_reader_c = combo.ExecuteReader();

                List<string> item = new List<string>();
                List<int> h_item = new List<int>();

                List<string> item1 = new List<string>();
                List<int> h_item1 = new List<int>();

                List<string> item2 = new List<string>();
                List<int> h_item2 = new List<int>();

                MySinger.Items.Clear();
                MyContinent.Items.Clear();
                comboBox3.Items.Clear();

                codes_of_singer.Clear();
                codes_of_continents.Clear();
                codes_of_country.Clear();
                while (que_reader_c.Read())
                {
                    item.Add(que_reader_c[4].ToString());
                    h_item.Add(Convert.ToInt32(que_reader_c[5]));
                    item1.Add(que_reader_c[2].ToString());
                    h_item1.Add(Convert.ToInt32(que_reader_c[3]));
                    item2.Add(que_reader_c[0].ToString());
                    h_item2.Add(Convert.ToInt32(que_reader_c[1]));


                   
                }
                que_reader_c.Close();

                object[] continents = item.OfType<string>().Distinct().ToArray();
                MyContinent.Items.AddRange(continents);

                int[] cont_code = h_item.OfType<int>().Distinct().ToArray();
                codes_of_continents.AddRange(cont_code);

                object[] country = item1.OfType<string>().Distinct().ToArray();
                comboBox3.Items.AddRange(country);

                int[] country_code = h_item1.OfType<int>().Distinct().ToArray();
                codes_of_country.AddRange(country_code);

                object[] singer = item2.OfType<string>().Distinct().ToArray();
                MySinger.Items.AddRange(singer);

                int[] singer_code = h_item2.OfType<int>().Distinct().ToArray();
                codes_of_singer.AddRange(singer_code);

                
            }
            if (this_list == 4) 
            {
                OleDbCommand combo = new OleDbCommand(query_for_change, myConnection);
                OleDbDataReader que_reader_c = combo.ExecuteReader();

                List<string> item = new List<string>();
                List<int> h_item = new List<int>();

                List<string> item1 = new List<string>();
                List<int> h_item1 = new List<int>();

                List<string> item2 = new List<string>();
                List<int> h_item2 = new List<int>();


                MyContinent.Items.Clear();
                MyCity.Items.Clear();
                comboBox3.Items.Clear();

                codes_of_continents.Clear();
                codes_of_city.Clear();
                codes_of_country.Clear();
                while (que_reader_c.Read())
                {
                    item.Add(que_reader_c[4].ToString());
                    h_item.Add(Convert.ToInt32(que_reader_c[5]));
                    item1.Add(que_reader_c[2].ToString());
                    h_item1.Add(Convert.ToInt32(que_reader_c[3]));
                    item2.Add(que_reader_c[0].ToString());
                    h_item2.Add(Convert.ToInt32(que_reader_c[1]));


                   
                }
                que_reader_c.Close();

                object[] continents = item.OfType<string>().Distinct().ToArray();
                MyContinent.Items.AddRange(continents);

                int[] cont_code = h_item.OfType<int>().Distinct().ToArray();
                codes_of_continents.AddRange(cont_code);

                object[] country = item1.OfType<string>().Distinct().ToArray();
                comboBox3.Items.AddRange(country);

                int[] country_code = h_item1.OfType<int>().Distinct().ToArray();
                codes_of_country.AddRange(country_code);

                object[] city = item2.OfType<string>().Distinct().ToArray();
                MyCity.Items.AddRange(city);

                int[] city_code = h_item2.OfType<int>().Distinct().ToArray();
                codes_of_city.AddRange(city_code);


               
            }   
        }
        private void MySinger_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            
            nGroup = codes_of_singer[MySinger.SelectedIndex];
            const_list = 4;
            string query = Get_Query(nGroup, nCity, nCountry, nContinent, const_list);
            Changing_Combo_Box(query, const_list);

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

            string sql2 = "SELECT " + singer_table + "," + city_table + "," + country_table + "," + continent_table;
            sql2 += table_query;
            OleDbCommand combo = new OleDbCommand(sql2, myConnection);
            OleDbDataReader que_reader_c = combo.ExecuteReader();
            while (que_reader_c.Read())
            {

                MyContinent.Items.Add(que_reader_c[6].ToString());
                comboBox3.Items.Add(que_reader_c[4].ToString());
                MyCity.Items.Add(que_reader_c[2].ToString());
                MySinger.Items.Add(que_reader_c[0].ToString());
                codes_of_continents.Add(Convert.ToInt32(que_reader_c[7]));
                codes_of_country.Add(Convert.ToInt32(que_reader_c[5]));
                codes_of_city.Add(Convert.ToInt32(que_reader_c[3]));
                codes_of_singer.Add(Convert.ToInt32(que_reader_c[1]));
            }
            que_reader_c.Close();
            Clear_Info();
            label5.Text = table.RowCount.ToString();
            label6.Text = MySinger.Items.Count.ToString();
            label7.Text = MyCity.Items.Count.ToString();
            label8.Text = comboBox3.Items.Count.ToString();
            label9.Text = MyContinent.Items.Count.ToString();
            button4.Enabled = true;
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

        private void MyContinent_TextChanged(object sender, EventArgs e)
        {
            
        }
    }
}
