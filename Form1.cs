//SET GLOBAL sql_mode=(SELECT REPLACE(@@sql_mode,'ONLY_FULL_GROUP_BY',''));

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using Newtonsoft.Json;
using System.Text.Unicode;
using System.Net.Http.Headers;

namespace Garant1._0
{
    public partial class Form1 : Form
    {
        MySqlCommand comm;

        int ID_User;

        bool find_row = false;

        string StartPath;

        string[] code_err_water;

        public Form1()
        {
            InitializeComponent();

            //code_err_water = new Dictionary<int, string>
            //{
            //    {71, "Не устанавливается" },
            //    {70, "Счётчик воды в сборе!" },
            //    {72, "Нарушена плома" },
            //    {74, "Предположительно обратная установка" },
            //    {80, "Паспорт!" },
            //    {82, "Не совпадает номер прибора с номером в паспорте" },
            //    {20, "Проливная часть в сборе!" },
            //    {21, "Крыльчатка не вращается" },
            //    {22, "Крыльчатка резко останавливается" },
            //    {23, "Всё вращается" },
            //    {25, "Отсутствует люфт без прокладки" },
            //    {26, "Другое" },
            //    {60, "Нарушение условий эксплуатации" },
            //    {62, "Превышение давления" },
            //    {63, "Механическое включение в камере" },
            //    {64, "Крыльчатка оплавлена" },
            //    {65, "Другое" },
            //    {66, "Известковый налет на всей поверхности камеры" },
            //    {30, "Корпус!" },
            //    {31, "Неперпендикулярная ось" },
            //    {32, "Налёт на оси" },
            //    {331, "Износ оси" },
            //    {34, "Занижена ось корпуса" },
            //    {35, "Завышена ось корпуса" },
            //    {37, "Течь корпуса (дефект литья)" },
            //    {38, "Течь из-под кольца" },
            //    {381, "Несоответствие диаметра (корпус)" },
            //    {382, "Несоответствие диаметра (крышка)" },
            //    {383, "Несоответствие размера 1,8" },
            //    {384, "Несоответствие размера 4,4" },
            //    {385, "Повреждение уплотнительного кольца" },
            //    {386, "Другое" },
            //    {39, "Несоответствующая резьба" },
            //    {11, "Некачественная сборка!" },
            //    {111, "Колёса другого типа" },
            //    {113, "Разгерметизация" },
            //    {114, "Повреждение деталей при сборке" },
            //    {12, "Интегратор!" },
            //    {121, "Не соответствует количество чёрных и красных барабанчиков" },
            //    {122, "Большой люфт" },
            //    {123, "Установка 1-4 разряда интегратора не на 0" },
            //    {13, "Стопорение!" },
            //    {131, "Мусор внутри СМ" },
            //    {132, "Некачественные барабанчики" },
            //    {133, "Некачественные колёса" },
            //    {134, "Другое" },
            //    {136, "Невозможно определить причину" },
            //    {137, "Сломана цапфа муфты" },
            //    {50, "Крыльчатка!" },
            //    {51, "Налипание на оси" },
            //    {52, "Налипание на магнитах" },
            //    {54, "Вылетел кольцевой магнит из крыльчатки" },
            //    {55, "Коррозия кольцевых магнитов" },
            //    {56, "Отсутствует магнит" },
            //    {40, "Крышка" },
            //    {42, "Занижен ⌀ 1,5 втулки" },
            //    {43, "Занижен наружный диаметр крышки" },
            //    {44, "Не запрессован подпишник" },
            //    {24, "Магниты в СМ и ПЧ разные" },
            //    {241, "Счётчик антимагнитный (СМ-2, ПЧ-6)" },
            //    {242, "Счётчик простой (СМ-2, ПЧ-6)" },
            //    {243, "Счётчик антимагнитный (СМ-6, ПЧ-2)" },
            //    {244, "Счётчик простой (СМ-6, ПЧ-2)" },
            //};
            code_err_water = new string[]
            {
                "70 - Счетчик воды в сборе!",
                "71 - не устанавливался",
                "72 - нарушена пломба",
                "74 - предположительно обратная установка",
                "80 - Паспорт!",
                "82 - не совпадает номер прибора с номером в паспорте",
                "20 - Проливная часть в сборе!",
                "21 - крыльчатка не вращается",
                "22 - крыльчатка резко останавливается",
                "23 - все вращается",
                "25 - отсутствует люфт без прокладки",
                "26 - другое",
                "60 - Нарушение условий эксплуатации!",
                "62 - превышение давления",
                "63 - механические включения в камере",
                "64 - крыльчатка оплавлена",
                "65 - другое",
                "66 - известковый налет на всей поверхности камеры",
                "30 - Корпус!",
                "31 - неперпендикулярная ось",
                "32 - налет на оси",
                "331 - износ оси",
                "34 - занижена ось корпуса",
                "35 - завышена ось корпуса",
                "37 - течь корпуса (дефект литья) ",
                "38 - течь из-под кольца",
                "381 - несоответствие диаметра (корпус)",
                "382 - несоответствие диаметра (крышка)",
                "383 - несоответствие размера 1,8",
                "384 - несоответствие размера 4,4",
                "385 - повреждение уплотнительного кольца",
                "386 - другое",
                "39 - несоответствующая резьба",
                "11 - Некачественная сборку!",
                "111 - Колеса другого типа",
                "113 - Разгерметизация",
                "114 - Повреждение деталей при сборке",
                "12 - Интегратор!",
                "121 - Не соответствует кол-во черных и красных барабанчиков",
                "122 - Большой люфт",
                "123 - Установка 1-4 разряда интегратора не на \"0\"",
                "13 - Стопорение!",
                "131 - Мусор внутри СМ",
                "132 - Некачественные барабанчики",
                "133 - Некачественные колеса",
                "134 - Другое",
                "136 - Невозможно определить причину",
                "137 - Сломана цапфа муфты",
                "50 - Крыльчатка!",
                "51 - Налипание на оси",
                "52 - Налипания на магнитах",
                "54 - Вылетел кольцевой магнит из крыльчатки",
                "55 - Коррозия кольцевых магнитов",
                "56 - Отсутствует магнит",
                "40 - Крышка!",
                "42 - Занижен ⌀ 1,5 втулки",
                "43 - Занижен наружный диаметр крышки",
                "44 - Не запрессован подшипник",
                "24 - Магниты в СМ и ПЧ разные",
                "241 - Счетчик антимагнитный (СМ-2, ПЧ-6)",
                "242 - Счетчик простой (СМ-2, ПЧ-6)",
                "243 - Счетчик антимагнитный (СМ-6, ПЧ-2)",
                "244 - Счетчик простой (СМ-6, ПЧ-2)"
            };

            //Form Menu = new Form();

            //MessageBox.Show(Application.StartupPath);
            StartPath = Application.StartupPath + "\\";

            comm = new MySqlCommand();
            //const string connect = "Database=Monitoring;Data Source=localhost;UserId=root;Password=;CharSet=utf8";
            const string connect = "server=127.0.0.1;user=root;database=Garant;port=3306;password=;charset=utf8mb4;convert zero datetime=True";
            try
            {
                MySqlConnection connection = new MySqlConnection(connect);
                comm.Connection = connection;
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                MessageBox.Show(ex.Message);
            }

            SetSQLMode();
            Find_users();
            //tabControl1.Enabled = false;

            checkedListBox2.Items.Clear();
            foreach (string code in code_err_water)
            {
                checkedListBox2.Items.Add(code);

            }

            /*
            MySqlDataReader reader = ExecutQuery("SELECT * FROM customers");
            if (reader != null)
                while (reader.Read())
                {
                    MessageBox.Show(reader["ID"].ToString());
                }
            comm.Connection.Close();
           
            //MessageBox.Show(DateTime.Now.ToString());

            //INSERT INTO `temptable`(`ID`, `Ser_Num`, `TypeMeter`, `Date`, `Solution`, `Codeb`) VALUES ([value-1],[value-2],[value-3],[value-4],[value-5],[value-6])

            MySqlDataReader reader = ExecutQuery("INSERT INTO  temptable (`Ser_Num
            , 'TypeMeter', 'Date', 'Solution', 'Codeb') VALUES ('46561984','СГВ-15','"+DateTime.Now.ToString()+"','Замена','1')");
            if (reader != null)
                while (reader.Read())
                {
                    //MessageBox.Show(reader["ID"].ToString());
                }
            comm.Connection.Close(); */
            //MessageBox.Show(DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"));

            //ExecutQuery_Insert("INSERT INTO  temptable (`Ser_Num`, `TypeMeter`, `Date`, `Solution`, `Codeb`) VALUES ('46561984','СГВ-15','" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "','Замена','1')");


            //btn_add_pribor.Enabled = false;
            //btn_refresh_pribor.Enabled = false;

            label24.Visible = false;
            label25.Visible = false;
            tb_num_pochta.Visible = false;
            tb_data_pochta.Visible = false;
            tb_data_pochta.Text = "";
            tb_num_pochta.Text = "";

            Create_DataGridView();
            fill_cb();
            FillDataGridViewCustomer();
            cb_kv_year_end.Text = DateTime.Now.ToString("yyyy");
            textBox1.Enabled = textBox2.Enabled = monthCalendar2.Enabled = monthCalendar1.Enabled = false;
        }

        void Find_users()
        {
            MySqlDataReader reader = ExecutQuery_Select("SELECT * FROM users");
            if (reader != null)
                while (reader.Read())
                {
                    comboBox_Users.Items.Add(reader["Descr"].ToString());
                }
            comm.Connection.Close();
        }
        
        //Решение ряда багов, связанных с типами данных MySQL версии > 8.0
        void SetSQLMode()
        {
            MySqlDataReader query1 = ExecutQuery_Select("SET GLOBAL sql_mode=(SELECT REPLACE(@@sql_mode,'ONLY_FULL_GROUP_BY',''));");
            comm.Connection.Close();

            //MySqlDataReader query3 = ExecutQuery_Select("SELECT @@GLOBAL.sql_mode global, @@SESSION.sql_mode session;SET sql_mode = '';SET GLOBAL sql_mode = '';");
            //comm.Connection.Close();
        }

        MySqlDataReader ExecutQuery_Select(String query)
        {
            //err_label.Text = "";
            MySqlDataReader reader = null;
            try
            {
                comm.CommandText = query;
                if (comm.Connection.State == ConnectionState.Closed)
                {
                    comm.Connection.Open();
                }
                reader = comm.ExecuteReader();


                //MessageBox.Show(comm.Connection.State.ToString());
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                //MessageBox.Show("err query = " + query);
                richTextBox1.Text += ("err query = " + query + "\nData\"" + ex.Message + "\"\n");
                MessageBox.Show(ex.Message);
                //reader = ExecutQuery_Select(query);
                /*tb_customer.Text = "";
                tb_date.Clear();
                tb_Serial_num.Clear();
                tb_type.Clear();
                cb_coded.Text = "";
                cb_solution.Text = "";*/
            }
            return reader;
        }

        int ExecutQuery_Insert(String query) // Вставска записей в БД
        {
            //err_label.Text = "";
            int i = 0;
            try
            {
                comm.CommandText = query;
                comm.Connection.Open();
                i = comm.ExecuteNonQuery();
                //richTextBox1.Text += ("ok insertquery = " + query + "\n");
            }
            catch (Exception e)
            {
                richTextBox1.Text += ("err query = " + query + "\nData\"" + e.Message + "\"\n");
            }
            comm.Connection.Close();
            return i;
        }

        private void comboBox_Users_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox_Users.Text.Trim() != "") {
                tabControl1.Enabled = true;
                MySqlDataReader reader = ExecutQuery_Select("SELECT * FROM users WHERE Descr = '" + comboBox_Users.Text + "'");
                if (reader != null)
                    while (reader.Read())
                    {
                        ID_User = Convert.ToInt32(reader["ID"].ToString());
                    }
                comm.Connection.Close();
            }
            else {
                tabControl1.Enabled = false;
            } 

            Num_Party.Items.Clear();
            cb_for_Otchet_Prih_Nak.Items.Clear();
            Refresh_DataGridView();

            MySqlDataReader reader2 = ExecutQuery_Select("SELECT DISTINCT IDParty FROM inwork");
            if (reader2 == null) return;
            if (reader2.HasRows != false) {

                while (reader2.Read()) {
                    Num_Party.Items.Add(reader2["IDParty"].ToString());
                    cb_for_Otchet_Prih_Nak.Items.Add(reader2["IDParty"].ToString());
                }
                //btn_add_pribor.Enabled = false;
                //btn_refresh_pribor.Enabled = true;
            }
            //Num_Party.Items.Add("Новая партия");
            comm.Connection.Close();

            Refresh_DataGridView();
        }

        private void Form1_Load(object sender, EventArgs e) { }

        private async void tb_Serial_num_TextChanged(object sender, EventArgs e)
        {
            string serial_num = tb_Serial_num.Text.Trim();
            tb_date.Text = "";
            tb_type.Text = "";
            cb_coded.Items.Clear();
            if (serial_num.Length == 8) {

                await Task.Delay(10);
                MySqlDataReader reader = ExecutQuery_Select("SELECT * FROM inwork WHERE Ser_Num = '" + serial_num + "'");
                if (reader == null) return;
                if (reader.HasRows != false) {

                    while (reader.Read()) {
                        //Num_Party.Text = reader["IDParty"].ToString();
                        tb_type.Text = reader["TypeMeter"].ToString();
                        tb_descr.Text = reader["Descr"].ToString();
                        tb_date.Text = reader.GetDateTime(7).ToString("yyyy-MM-dd hh:mm:ss");
                        //tb_date.Text = reader["DateCreate"].ToString();
                        //tb_dateAnaliz.Text = reader["DateAnaliz"].ToString();
                        tb_dateAnaliz.Text = reader.GetDateTime(12).ToString("yyyy-MM-dd hh:mm:ss");
                        //cb_coded.Text = (reader["Codeb"].ToString().Trim() != "") ? reader["Codeb"].ToString() : "";

                        cb_solution.Text = reader["Solution"].ToString();
                        tb_customer.Text = reader["CustomerID"].ToString();
                        cb_sposob_dost.Text = reader["sposob_dost"].ToString();
                        tb_num_pochta.Text = (DBNull.Value.Equals(reader["num_dost"])) ? "" : reader["num_dost"].ToString();
                        tb_data_pochta.Text = (DBNull.Value.Equals(reader["num_dost"])) ? "" : reader.GetDateTime(17).ToString("yyyy-MM-dd hh:mm:ss");
                        tb_narab.Text = reader["narabotka"].ToString();

                        radioButton1.Checked = (reader["kit"].ToString() == "1") ? true : false;
                        radioButton2.Checked = (reader["kit"].ToString() == "0") ? true : false;

                        cb_coded.Items.AddRange(reader["Codeb"].ToString().Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries));



                        btn_add_pribor.Text = "Обновить";

                    }
                    //btn_add_pribor.Enabled = false;
                    //btn_refresh_pribor.Enabled = true;
                } else {
                    btn_add_pribor.Text = "Добавить новый";
                    comm.Connection.Close();

                    reader = ExecutQuery_Select("SELECT * FROM temptable WHERE Ser_Num = '" + serial_num + "'");
                    if (reader != null) {

                        while (reader.Read()) {
                            //MessageBox.Show(reader["Date"].ToString());
                            tb_type.Text = reader["TypeMeter"].ToString();
                            tb_date.Text = reader.GetDateTime(3).ToString("yyyy-MM-dd hh:mm:ss");
                            //tb_date.Text = reader["Date"].ToString();
                            cb_coded.Text = reader["Codeb"].ToString();
                            //cb_solution.Text = reader["Solution"].ToString();
                            //tb_customer.Text = reader["CustomerID"].ToString();
                            tb_dateAnaliz.Text = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");
                        }
                    }
                    //btn_add_pribor.Enabled = true;
                    //btn_refresh_pribor.Enabled = false;
                }
                comm.Connection.Close();
            }
        }

        private void btn_add_pribor_Click(object sender, EventArgs e)
        {
            string codes_braka = "";
            foreach (string d in cb_coded.Items)
            {
                codes_braka += d + '|';
            }
            MessageBox.Show(codes_braka);
            if (btn_add_pribor.Text.Trim() == "Обновить")
            {
                find_row = false;
                int kit = (radioButton1.Checked == true) ? 1 : 0;
                //UPDATE `inwork` SET `ID`=[value-1],`IDParty`=[value-2],`Ser_Num`=[value-3],`TypeMeter`=[value-4],`DateCreate`=[value-5],`Solution`=[value-6],`Codeb`=[value-7],`CustomerID`=[value-8],`UserID`=[value-9],`DateCheck`=[value-10] WHERE 1
                int res = ExecutQuery_Insert("UPDATE `inwork` SET `Ser_Num`='"
                    + tb_Serial_num.Text + "',`TypeMeter`='"
                    + tb_type.Text + "',`DateCreate`='" + tb_date.Text + "',`Solution`='"
                    + cb_solution.Text + "',`Codeb`='" + codes_braka + "',`Code_err_user`='" + cb_code_cust.Text + "',`CustomerID`='" + tb_customer.Text + "',`UserID`='"
                    + ID_User + "',`DateAnaliz`='" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "',`Descr`='" + tb_descr.Text + "',`kit`='"
                    + kit + "',`sposob_dost`='" + cb_sposob_dost.Text + "',`num_dost`='" + tb_num_pochta.Text + "',`date_dost`='" + tb_data_pochta.Text
                    + "',`narabotka`='" + tb_narab.Text + "' WHERE Ser_Num='" + tb_Serial_num.Text + "';");

                tb_customer.Text = "";
                tb_date.Clear();
                tb_Serial_num.Clear();
                tb_type.Clear();
                cb_coded.Items.Clear();
                //cb_solution.Text = "";
                Refresh_DataGridView();
                tb_narab.Text = "";

            } else {

                if (Num_Party.Text == "Новая партия") {
                    MySqlDataReader reader = ExecutQuery_Select("SELECT MAX(IDParty) as IDParty FROM inwork");
                    if (reader != null)

                        while (reader.Read())
                        {
                            Num_Party.Items[Num_Party.Items.Count - 1] = Convert.ToInt32(reader["IDParty"]) + 1;
                            cb_for_Otchet_Prih_Nak.Items.Add(Convert.ToInt32(reader["IDParty"]) + 1);
                            Num_Party.Items.Add("Новая партия");
                        }
                    comm.Connection.Close();
                }

                find_row = false;
                int kit = (radioButton1.Checked == true) ? 1 : 0;
                //int num_dost = (tb_num_pochta.Text == "") ? 0 : Convert.ToInt32(tb_num_pochta.Text);
                //string date_dost = (tb_data_pochta.Text == "") ? "" : tb_data_pochta.Text;

                if (cb_sposob_dost.Text == "Непосредственно от потреб.") { 
                    int res = ExecutQuery_Insert("INSERT INTO `inwork` (`ID`, `IDParty`, `Ser_Num`, `TypeMeter`, `DateCreate`, `Solution`, `Codeb`, `CustomerID`, `UserID`, `DatePriem`, `DateAnaliz`, `Descr`, `kit`, `sposob_dost`, `num_dost`, `date_dost`, `narabotka`,`Code_err_user`) VALUES (NULL, '"
                    + Num_Party.Text + "', '" + tb_Serial_num.Text + "', '" + tb_type.Text + "', '"
                    + tb_date.Text + "', '" + cb_solution.Text + "', '" + codes_braka + "', '" + tb_customer.Text
                    + "', '" + ID_User + "', '" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "', '"
                    + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "', '" + tb_descr.Text + "', '"
                    + kit + "', '"
                    + cb_sposob_dost.Text + "', "
                    + "NULL, "
                    + "NULL, '"
                    + tb_narab.Text + "', '"
                    + cb_code_cust.Text + "');"); 
                } else {
                    int res = ExecutQuery_Insert("INSERT INTO `inwork` (`ID`, `IDParty`, `Ser_Num`, `TypeMeter`, `DateCreate`, `Solution`, `Codeb`, `CustomerID`, `UserID`, `DatePriem`, `DateAnaliz`, `Descr`, `kit`, `sposob_dost`, `num_dost`, `date_dost`, `narabotka`,`Code_err_user`) VALUES (NULL, '"
                    + Num_Party.Text + "', '" + tb_Serial_num.Text + "', '" + tb_type.Text + "', '"
                    + tb_date.Text + "', '" + cb_solution.Text + "', '" + codes_braka + "', '" + tb_customer.Text
                    + "', '" + ID_User + "', '" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "', '"
                    + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "', '" + tb_descr.Text + "', '"
                    + kit + "', '"
                    + cb_sposob_dost.Text + "', '"
                    + tb_num_pochta.Text + "', '"
                    + tb_data_pochta.Text + "', '"
                    + tb_narab.Text + "', '"
                    + cb_code_cust.Text + "');");
                }
                //int res = ExecutQuery_Insert("INSERT INTO `inwork` (`ID`, `IDParty`, `Ser_Num`, `TypeMeter`, `DateCreate`, `Solution`, `CustomerID`, `UserID`, `DatePriem`, `DateAnaliz`, `Descr`) VALUES (NULL, '" + Num_Party.Text + "', '" + tb_Serial_num.Text + "', '" + tb_type.Text + "', '" + tb_date.Text + "', '" + cb_solution.Text + "', '" + tb_customer.Text + "', '" + ID_User + "', '" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "', '" + tb_dateAnaliz.Text + "', '" + tb_descr.Text + "');");
                //MessageBox.Show(res + "");
                //tb_customer.Text = "";
                tb_date.Clear();
                tb_Serial_num.Clear();
                tb_type.Clear();
                cb_coded.Items.Clear();
                //cb_solution.Text = "";
                tb_dateAnaliz.Text = "";
                tb_narab.Text = "";
                Refresh_DataGridView();
            }
        }

        private void btn_refresh_pribor_Click(object sender, EventArgs e)
        {
            find_row = false;

            int code = 0;
            try
            {
                code = Convert.ToInt32(cb_coded.Text);
            }
            catch { }
            int kit = (radioButton1.Checked == true) ? 1 : 0;
            //UPDATE `inwork` SET `ID`=[value-1],`IDParty`=[value-2],`Ser_Num`=[value-3],`TypeMeter`=[value-4],`DateCreate`=[value-5],`Solution`=[value-6],`Codeb`=[value-7],`CustomerID`=[value-8],`UserID`=[value-9],`DateCheck`=[value-10] WHERE 1
            int res = ExecutQuery_Insert("UPDATE `inwork` SET `Ser_Num`='" + tb_Serial_num.Text + "',`TypeMeter`='" + tb_type.Text + "',`DateCreate`='" + tb_date.Text + "',`Solution`='" + cb_solution.Text + "',`Codeb`='" + code + "',`CustomerID`='" + tb_customer.Text + "',`UserID`='" + ID_User + "',`DateAnaliz`='" + tb_dateAnaliz.Text + "',`Descr`='" + tb_descr.Text + "',`kit`='" + kit + "' WHERE Ser_Num='" + tb_Serial_num.Text + "';");

            tb_customer.Text = "";
            tb_date.Clear();
            tb_Serial_num.Clear();
            tb_type.Clear();
            cb_coded.Text = "";
            cb_solution.Text = "";
            Refresh_DataGridView();

        }

        void Create_DataGridView()
        {
            var con_num = new DataGridViewColumn();
            con_num.HeaderText = "№";
            //con_num.Width = 25; //ширина колонки
            con_num.ReadOnly = true; //значение в этой колонке нельзя править
            con_num.Name = "Number";
            //con_num.Frozen = true; //флаг, что данная колонка всегда отображается на своем месте
            con_num.CellTemplate = new DataGridViewTextBoxCell();

            var column1 = new DataGridViewColumn();
            column1.HeaderText = "Серийный номер"; //текст в шапке
            column1.Name = "SerNum"; //текстовое имя колонки, его можно использовать вместо обращений по индексу
            column1.CellTemplate = new DataGridViewTextBoxCell(); //тип нашей колонки

            var column2 = new DataGridViewColumn();
            column2.HeaderText = "Тип";
            column2.Name = "Type";
            column2.CellTemplate = new DataGridViewTextBoxCell();

            var column3 = new DataGridViewColumn();
            column3.HeaderText = "Дата производства";
            column3.Name = "Date_made";
            //column3.Width = 100;
            column3.CellTemplate = new DataGridViewTextBoxCell();

            var column4 = new DataGridViewColumn();
            column4.HeaderText = "Потребитель";
            column4.Name = "Customer";
            column4.CellTemplate = new DataGridViewTextBoxCell();

            var column5 = new DataGridViewColumn();
            column5.HeaderText = "Причина";
            column5.Name = "Solution";
            //column5.Width = 100;
            column5.CellTemplate = new DataGridViewTextBoxCell();

            var column5_5 = new DataGridViewColumn();
            column5_5.HeaderText = "Код (Потреб.)";
            column5_5.Name = "CodeA";
            //column6.Width = 50;
            column5_5.CellTemplate = new DataGridViewTextBoxCell();

            var column6 = new DataGridViewColumn();
            column6.HeaderText = "Код дефекта";
            column6.Name = "CodeB";
            //column6.Width = 50;
            column6.CellTemplate = new DataGridViewTextBoxCell();

            var column7 = new DataGridViewColumn();
            column7.HeaderText = "Анализ проведен";
            column7.Name = "User";
            column7.CellTemplate = new DataGridViewTextBoxCell();

            /*var column8 = new DataGridViewColumn();
            column8.HeaderText = "Дата приема";
            column8.Name = "Date_priem";
            //column8.Width = 100;
            column8.CellTemplate = new DataGridViewTextBoxCell();*/
            var column8 = new DataGridViewColumn();
            column8.HeaderText = "Наработка";
            column8.Name = "Narabotka";
            //column8.Width = 100;
            column8.CellTemplate = new DataGridViewTextBoxCell();

            var column9 = new DataGridViewColumn();
            column9.HeaderText = "Дата анализа";
            column9.Name = "Date_Analiz";
            //column9.Width = 100;
            column9.CellTemplate = new DataGridViewTextBoxCell();


            var column10 = new DataGridViewColumn();
            column10.HeaderText = "Примечание";
            column10.Name = "Descr";
            //column10.Width = 150;
            column10.CellTemplate = new DataGridViewTextBoxCell();

            var column11 = new DataGridViewColumn();
            column11.HeaderText = "Комплект";
            column11.Name = "Kit";
            //column10.Width = 150;
            column11.CellTemplate = new DataGridViewTextBoxCell();

            dataGridView1.Columns.Add(con_num);
            dataGridView1.Columns.Add(column8);
            dataGridView1.Columns.Add(column2);
            dataGridView1.Columns.Add(column1);
            dataGridView1.Columns.Add(column3);
            dataGridView1.Columns.Add(column4);
            dataGridView1.Columns.Add(column5);
            dataGridView1.Columns.Add(column5_5);
            dataGridView1.Columns.Add(column6);
            dataGridView1.Columns.Add(column7);
            dataGridView1.Columns.Add(column9);
            dataGridView1.Columns.Add(column10);
            dataGridView1.Columns.Add(column11);


            dataGridView1.AllowUserToAddRows = false; //запрешаем пользователю самому добавлять строки
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.AllowUserToOrderColumns = false;


            //Refresh_DataGridView();
        }
        void Refresh_DataGridView()
        {
            dataGridView1.Rows.Clear();
            MySqlDataReader reader = ExecutQuery_Select("SELECT * FROM `inwork` WHERE IDParty = '" + Num_Party.Text + "'");
            int i = 1;
            if (reader == null) return;
            if (reader.HasRows != false)
            {
                while (reader.Read())
                {
                    dataGridView1.Rows.Add(i, reader["narabotka"].ToString(), reader["TypeMeter"].ToString(), reader["Ser_Num"].ToString(), reader["DateCreate"].ToString(),
                        reader["CustomerID"].ToString(), reader["Solution"].ToString(), reader["Code_err_user"].ToString(), reader["Codeb"].ToString(), reader["UserID"].ToString(), reader["DateAnaliz"].ToString(), reader["Descr"].ToString(), reader["kit"].ToString());
                    i++;
                }
                //btn_add_pribor.Enabled = false;
                //btn_refresh_pribor.Enabled = true;
            }
            comm.Connection.Close();
            find_row = true;


            tb_customer.Items.Clear();
            reader = ExecutQuery_Select("SELECT * FROM `customers`");
            if (reader.HasRows != false)
            {
                while (reader.Read())
                {
                    tb_customer.Items.Add(reader["Descr"].ToString());
                }
                //btn_add_pribor.Enabled = false;
                //btn_refresh_pribor.Enabled = true;
            }
            comm.Connection.Close();

        }

        private void dataGridView1_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (find_row == false)
                return;

            //MessageBox.Show(dataGridView1["SerNum", e.RowIndex].Value.ToString());
            //MessageBox.Show("RowEnter");
            tb_Serial_num.Text = dataGridView1["SerNum", e.RowIndex].Value.ToString();
        }

        private void Num_Party_SelectedIndexChanged(object sender, EventArgs e)
        {
            Refresh_DataGridView();
        }

        private void Num_Party_TextChanged(object sender, EventArgs e)
        {
            //Refresh_DataGridView();
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void err_label_Click(object sender, EventArgs e)
        {
            //err_label.Text = "";
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {

        }

        private void Create_Priz_Nak_Click(object sender, EventArgs e)
        {
            if (cb_for_Otchet_Prih_Nak.Text.Trim() == "") return;

            MySqlDataReader reader1 = ExecutQuery_Select("SELECT * FROM inwork WHERE IDParty = '" + cb_for_Otchet_Prih_Nak.Text + "'");
            if (reader1 != null)
            {
                if (!reader1.HasRows)
                {
                    MessageBox.Show("Приходная накладная не найдена.");
                    comm.Connection.Close();
                    return;
                }
            }
            comm.Connection.Close();
            Word.Application appWord;
            Word.Document docWord = null;
            object missobj = System.Reflection.Missing.Value;
            object falseobj = false;
            object trueobj = true;

            appWord = new Word.Application();
            object path_sh = StartPath + "Shablon_Prih_Nak.docx";
            try
            {
                docWord = appWord.Documents.Add(ref path_sh, ref missobj, ref missobj, ref missobj);
            }
            catch (Exception err)
            {
                docWord.Close(ref falseobj, ref missobj, ref missobj);
                appWord.Quit(ref missobj, ref missobj, ref missobj);
                docWord = null;
                appWord = null;
                throw err;
            }

            object refNum = "NUM";
            object refDate = "Date";
            object refCustomer = "Customer";
            object all_kit = "all_kit";
            object all_kol = "all_kol";
            object refUSER = "USER";

            object refNum2 = "NUM2";
            object refDate2 = "Date2";
            object refCustomer2 = "Customer2";
            object all_kit2 = "all_kit2";
            object all_kol2 = "all_kol2";
            object refUSER2 = "USER2";

            Word.Bookmark bookmark_NUM = docWord.Bookmarks[ref refNum];
            Word.Bookmark bookmark_Date = docWord.Bookmarks[ref refDate];
            Word.Bookmark bookmark_Customer = docWord.Bookmarks[ref refCustomer];
            Word.Bookmark bookmark_USER = docWord.Bookmarks[ref refUSER];

            Word.Bookmark bookmark_NUM2 = docWord.Bookmarks[ref refNum2];
            Word.Bookmark bookmark_Date2 = docWord.Bookmarks[ref refDate2];
            Word.Bookmark bookmark_Customer2 = docWord.Bookmarks[ref refCustomer2];
            Word.Bookmark bookmark_USER2 = docWord.Bookmarks[ref refUSER2];

            bookmark_NUM.Range.Text = cb_for_Otchet_Prih_Nak.Text;
            bookmark_Date.Range.Text = DateTime.Now.ToString("D");

            bookmark_NUM2.Range.Text = cb_for_Otchet_Prih_Nak.Text;
            bookmark_Date2.Range.Text = DateTime.Now.ToString("D");

            //CustomerID
            string USER_ID = "0";
            MySqlDataReader reader = ExecutQuery_Select("SELECT CustomerID, UserID FROM inwork WHERE IDParty = '" + cb_for_Otchet_Prih_Nak.Text + "'");
            if (reader != null)
            {
                while (reader.Read())
                {
                    bookmark_Customer.Range.Text = reader["CustomerID"].ToString();
                    bookmark_Customer2.Range.Text = reader["CustomerID"].ToString();

                    USER_ID = reader["UserID"].ToString();
                    break;
                }
            }
            comm.Connection.Close();

            //USERNAME
            MySqlDataReader User = ExecutQuery_Select("SELECT Descr FROM users WHERE ID = '" + USER_ID + "'");
            if (User != null)
            {
                while (User.Read())
                {
                    bookmark_USER.Range.Text = User["Descr"].ToString();
                    bookmark_USER2.Range.Text = User["Descr"].ToString();
                    break;
                }
            }
            comm.Connection.Close();

            Word.Table tableWord = docWord.Tables[1].Tables[2];
            Word.Table tableWord2 = docWord.Tables[1].Tables[6];



            int row = 2;

            int all_kol_int = 0;
            int all_kit_int = 0;


            List<string> types_meter = new List<string>();
            List<string> reasons_meter = new List<string>();

            MySqlDataReader type_data = ExecutQuery_Select("SELECT DISTINCT TypeMeter FROM inwork WHERE IDParty = '" + cb_for_Otchet_Prih_Nak.Text + "'");
            {
                if (type_data != null)
                {
                    while (type_data.Read())
                    {
                        types_meter.Add(type_data["TypeMeter"].ToString());
                        //MessageBox.Show(type_data["TypeMeter"].ToString());
                    }
                }
            }
            comm.Connection.Close();

            foreach (string type in types_meter)
            {
                reasons_meter.Clear();
                MySqlDataReader reason = ExecutQuery_Select("SELECT DISTINCT Solution FROM inwork WHERE IDParty = '" + cb_for_Otchet_Prih_Nak.Text + "' AND TypeMeter = '" + type + "'");
                {
                    if (reason != null)
                    {
                        while (reason.Read())
                        {
                            reasons_meter.Add(reason["Solution"].ToString());
                            //MessageBox.Show(reason["Solution"].ToString());

                        }
                    }
                }
                comm.Connection.Close();
                //MessageBox.Show(reasons_meter.Count.ToString());

                foreach (string solution in reasons_meter)
                {
                    //MessageBox.Show(type + "  " + solution);
                    MySqlDataReader withoutKit = ExecutQuery_Select("SELECT COUNT(*) as count FROM inwork WHERE IDParty = '" + cb_for_Otchet_Prih_Nak.Text + "' AND TypeMeter = '" + type + "' AND Solution = '" + solution + "'");
                    if (withoutKit != null)
                    {
                        while (withoutKit.Read())
                        {
                            if (row != 2)
                            {
                                tableWord.Rows.Add(ref missobj);
                                tableWord2.Rows.Add(ref missobj);

                            }
                            tableWord.Cell(row, 1).Range.Text = type;
                            tableWord.Cell(row, 2).Range.Text = withoutKit["count"].ToString();
                            tableWord.Cell(row, 3).Range.Text = "0";
                            tableWord.Cell(row, 4).Range.Text = solution;

                            tableWord2.Cell(row, 1).Range.Text = type;
                            tableWord2.Cell(row, 2).Range.Text = withoutKit["count"].ToString();
                            tableWord2.Cell(row, 3).Range.Text = "0";
                            tableWord2.Cell(row, 4).Range.Text = solution;

                            all_kol_int += Convert.ToInt32(withoutKit["count"]);
                            row++;
                            //MessageBox.Show(type + "  " + solution + "  " + withoutKit["count"]);
                        }
                    }
                    comm.Connection.Close();
                    MySqlDataReader withKit = ExecutQuery_Select("SELECT COUNT(*) as count FROM inwork WHERE IDParty = '" + cb_for_Otchet_Prih_Nak.Text + "' AND TypeMeter = '" + type + "' AND Solution = '" + solution + "' AND kit = '1'");
                    if (withKit != null)
                    {
                        while (withKit.Read())
                        {
                            //if (row != 2) tableWord.Rows.Add(tableWord.Rows[2]);
                            tableWord.Cell(row - 1, 3).Range.Text = withKit["count"].ToString();
                            tableWord2.Cell(row - 1, 3).Range.Text = withKit["count"].ToString();

                            all_kit_int += Convert.ToInt32(withKit["count"]);
                            //MessageBox.Show(type + "  " + solution + "  " + withoutKit["count"]);
                        }
                    }
                    comm.Connection.Close();
                }


            }


            Word.Bookmark bookmark_all_kol = docWord.Bookmarks[ref all_kol];
            Word.Bookmark bookmark_all_kit = docWord.Bookmarks[ref all_kit];

            bookmark_all_kol.Range.Text = all_kol_int.ToString();
            bookmark_all_kit.Range.Text = all_kit_int.ToString();

            Word.Bookmark bookmark_all_kol2 = docWord.Bookmarks[ref all_kol2];
            Word.Bookmark bookmark_all_kit2 = docWord.Bookmarks[ref all_kit2];

            bookmark_all_kol2.Range.Text = all_kol_int.ToString();
            bookmark_all_kit2.Range.Text = all_kit_int.ToString();

            appWord.Visible = true;


            MySqlDataReader acts = ExecutQuery_Select("SELECT COUNT(*) as count FROM acts WHERE IDParty = '" + cb_for_Otchet_Prih_Nak.Text + "'");
            if (acts != null)
            {
                while (acts.Read())
                {
                    if (Convert.ToInt32(acts["count"]) > 0)
                    {
                        comm.Connection.Close();
                        //MessageBox.Show("Есть строки");

                    }
                    else
                    {
                        //MessageBox.Show("нет строки");
                        comm.Connection.Close();
                        Form_acts(cb_for_Otchet_Prih_Nak.Text);

                    }
                    break;
                }
            }
            else
            {
                MessageBox.Show("Записей нет");
            }
            comm.Connection.Close();
            if (checkBox1.Checked)
            {
                Show_All_Acts(Convert.ToInt32(cb_for_Otchet_Prih_Nak.Text));
            }
        }
        void fill_cb()
        {
            MySqlDataReader reader = ExecutQuery_Select("SELECT * FROM reasonreturn");
            if (reader != null)
            {
                while (reader.Read())
                {
                    cb_solution.Items.Add(reader["Reason"].ToString());
                }
            }
            comm.Connection.Close();
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {

        }

        void FillDataGridViewCustomer()
        {
            dataGridView2.Rows.Clear();
            MySqlDataReader reader = ExecutQuery_Select("SELECT * FROM `customers`");
            int i = 1;
            if (reader == null) return;
            if (reader.HasRows != false)
            {
                while (reader.Read())
                {
                    dataGridView2.Rows.Add(reader["ID"].ToString(), reader["Descr"].ToString(), reader["ContFace"].ToString(), reader["Phone"].ToString(), reader["_Index"].ToString(),
                        reader["Resp"].ToString(), reader["Oblast"].ToString(), reader["City"].ToString(), reader["Street"].ToString(), reader["Num_h"].ToString(), reader["Num_f"].ToString());
                    i++;
                }
            }
            comm.Connection.Close();
        }

        private void cust_id_TextChanged(object sender, EventArgs e)
        {
            add_customer.Text = (cust_id.Text.Trim() == "") ? "Добавить нового" : "Обновить";

            if (cust_id.Text.Trim() != "")
            {
                MySqlDataReader reader = ExecutQuery_Select("SELECT * FROM `customers` WHERE ID = '" + cust_id.Text.Trim() + "'");
                //int i = 1;
                if (reader == null) return;
                if (reader.HasRows != false)
                {
                    while (reader.Read())
                    {
                        cust_name.Text = reader["Descr"].ToString();
                        cust_face.Text = reader["ContFace"].ToString();
                        cust_phone.Text = reader["Phone"].ToString();
                        cust_index.Text = reader["_Index"].ToString();
                        cust_resp.Text = reader["Resp"].ToString();
                        cust_raion.Text = reader["Oblast"].ToString();
                        cust_city.Text = reader["City"].ToString();
                        cust_street.Text = reader["Street"].ToString();
                        cust_house.Text = reader["Num_h"].ToString();
                        cust_flat.Text = reader["Num_f"].ToString();
                    }
                }
                comm.Connection.Close();
            }
        }

        private void add_customer_Click(object sender, EventArgs e)
        {
            if (cust_id.Text.Trim() != "")
            {
                int res = ExecutQuery_Insert("UPDATE `customers` SET `Descr`='" + cust_name.Text + "" +
                    "',`ContFace`='" + cust_face.Text + "" +
                    "',`Phone`='" + cust_phone.Text + "',`_Index`='" + cust_index.Text +
                    "',`Resp`='" + cust_resp.Text + "',`Oblast`='" + cust_raion.Text +
                    "',`City`='" + cust_city.Text + "',`Street`='" + cust_street.Text +
                    "',`Num_h`='" + cust_house.Text + "',`Num_f`='" + cust_flat.Text + "' WHERE ID='" + cust_id.Text + "';");
            }
            else
            {
                int res = ExecutQuery_Insert("INSERT INTO `customers` (`ID`, `Descr`, `ContFace`, `Phone`, `_Index`, `Resp`, `Oblast`, `City`, `Street`, `Num_h`, `Num_f`) " +
                    "VALUES (NULL, '" + cust_name.Text + "', '" + cust_face.Text + "', '" + cust_phone.Text + "', '" + cust_index.Text + "', '"
                    + cust_resp.Text + "', '" + cust_raion.Text + "', '" + cust_city.Text + "', '" + cust_street.Text + "', '" + cust_house.Text + "', '" + cust_flat.Text + "');");
            }

            FillDataGridViewCustomer();
            Refresh_DataGridView();
        }

        private void dataGridView2_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            cust_id.Text = dataGridView2["Column1", e.RowIndex].Value.ToString();
        }

        private void sbros_customer_Click(object sender, EventArgs e)
        {
            cust_id.Text = "";
            cust_name.Text = "";
            cust_face.Text = "";
            cust_phone.Text = "";
            cust_index.Text = "";
            cust_resp.Text = "";
            cust_raion.Text = "";
            cust_city.Text = "";
            cust_street.Text = "";
            cust_house.Text = "";
            cust_flat.Text = "";
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_sposob_dost.Text == "Непосредственно от потреб.")
            {
                label24.Visible = false;
                label25.Visible = false;
                tb_num_pochta.Visible = false;
                tb_data_pochta.Visible = false;
                tb_data_pochta.Text = "";
                tb_num_pochta.Text = "";
            }
            else
            {
                label24.Visible = true;
                label25.Visible = true;
                tb_num_pochta.Visible = true;
                tb_data_pochta.Visible = true;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Show_Act(1);

            /*meter meter1 = new meter() { Descr = "Лалка" };
            string json = JsonSerializer.Serialize<meter>(meter1);

            //act act1 = new act { num = 1, schts = json };
            //json = JsonSerializer.Serialize<meter[]>(meters);
            MessageBox.Show(json);
            meter met2 = JsonSerializer.Deserialize<meter>(json);
            //meter[] m2 = JsonSerializer.Deserialize<meter[]>();
            MessageBox.Show(met2.Descr);*/

            string query = "SELECT * FROM inwork WHERE (";
            string text_data = "Отчет по \"";
            bool have_meter = false;
            bool have_code_err = false;
            string[] errors = new string[checkedListBox2.Items.Count];
            int error_index = 0;

            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (checkedListBox1.GetItemChecked(i))
                {
                    if (have_meter == true)
                    {
                        query += " OR ";
                    }
                    have_meter = true;
                    query += "TypeMeter = '" + checkedListBox1.Items[i] + "'";
                    text_data += checkedListBox1.Items[i] + " ";
                }
            }
            if (have_meter == false)
            {
                query += "1";
            }
            query += ") AND (";
            text_data += "\" по \"";
            have_meter = false;
            for (int i = 0; i < checkedListBox2.Items.Count; i++)
            {
                if (checkedListBox2.GetItemChecked(i))
                {
                    if (have_meter == true)
                    {
                        query += " OR ";
                    }
                    have_meter = true;
                    query += "Codeb LIKE '%" + checkedListBox2.Items[i] + "%'";
                    text_data += checkedListBox2.Items[i] + " ";
                    have_code_err = true;

                    errors[error_index] = checkedListBox2.Items[i] + "";
                    error_index++;
                }
            }
            if (have_meter == false)
            {
                query += "1";
            }
            query += ") AND (DatePriem BETWEEN ";
            text_data += "\" за ";

            if (checkBox2.Checked == true)
            {
                text_data += cb_kv_year_start.Text + "|" + cb_kv_start.Text + "| : " + cb_kv_year_end.Text + "|" + cb_kv_end.Text + "|";
                if (cb_kv_start.Text == cb_kv_end.Text && cb_kv_year_end.Text == cb_kv_year_start.Text)
                {
                    switch (cb_kv_start.Text)
                    {
                        case "1": query += "'" + cb_kv_year_start.Text + "-" + "01" + "-" + "01 00:00:00' and '" + cb_kv_year_start.Text + "-" + "03" + "-" + "31 23:59:59'"; break;
                        case "2": query += "'" + cb_kv_year_start.Text + "-" + "04" + "-" + "01 00:00:00' and '" + cb_kv_year_start.Text + "-" + "06" + "-" + "30 23:59:59'"; break;
                        case "3": query += "'" + cb_kv_year_start.Text + "-" + "07" + "-" + "01 00:00:00' and '" + cb_kv_year_start.Text + "-" + "09" + "-" + "30 23:59:59'"; break;
                        case "4": query += "'" + cb_kv_year_start.Text + "-" + "10" + "-" + "01 00:00:00' and '" + cb_kv_year_start.Text + "-" + "12" + "-" + "31 23:59:59'"; break;
                        default: query += "'" + cb_kv_year_start.Text + "-" + "01" + "-" + "01 00:00:00' and '" + cb_kv_year_start.Text + "-" + "03" + "-" + "31 23:59:59'"; break;
                    }
                }
                else
                {
                    switch (cb_kv_start.Text)
                    {
                        case "1": query += "'" + cb_kv_year_start.Text + "-" + "01" + "-" + "01 00:00:00'"; break;
                        case "2": query += "'" + cb_kv_year_start.Text + "-" + "04" + "-" + "01 00:00:00'"; break;
                        case "3": query += "'" + cb_kv_year_start.Text + "-" + "07" + "-" + "01 00:00:00'"; break;
                        case "4": query += "'" + cb_kv_year_start.Text + "-" + "10" + "-" + "01 00:00:00'"; break;
                        default: query += "'" + cb_kv_year_start.Text + "-" + "01" + "-" + "01 00:00:00'"; break;
                    }
                    query += " and ";
                    switch (cb_kv_end.Text)
                    {
                        case "1": query += "'" + cb_kv_year_end.Text + "-" + "03" + "-" + "31 23:59:59'"; break;
                        case "2": query += "'" + cb_kv_year_end.Text + "-" + "06" + "-" + "30 23:59:59'"; break;
                        case "3": query += "'" + cb_kv_year_end.Text + "-" + "09" + "-" + "30 23:59:59'"; break;
                        case "4": query += "'" + cb_kv_year_end.Text + "-" + "12" + "-" + "31 23:59:59'"; break;
                        default: query += "'" + cb_kv_year_end.Text + "-" + "03" + "-" + "31 23:59:59'"; break;
                    }
                }
            }
            else
            {
                text_data += textBox1.Text + " : " + textBox2.Text;
                query += "'" + textBox1.Text + " 00:00:00' and '" + textBox2.Text + " 23:59:59'";
            }

            query += ")";


            richTextBox1.Text = query;

            string[] rimsk = new string[] { "0", "I", "II", "III", "IV" };

            MySqlDataReader data = ExecutQuery_Select(query);
            //var dataList = new List<string>();

            //if (data != null)
            //{
            //    MessageBox.Show("Not empty query");
            //    while (data.Read())
            //    {
            //        for (int index = 0; index < data.FieldCount; index++)
            //        {
            //            dataList.Add(data[index].ToString());
            //        }
            //    }
            //    data.Close();
            //}

            //var jsonData = JsonConvert.SerializeObject(dataList);
            //MessageBox.Show(jsonData);


            if (data != null)
            {
                Excel.Application ex = new Microsoft.Office.Interop.Excel.Application();
                ex.Visible = true;
                ex.SheetsInNewWorkbook = 1;
                Excel.Workbook workBook = ex.Workbooks.Add(Type.Missing);
                ex.DisplayAlerts = false;
                Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets[1];
                sheet.Name = "Отчет " + DateTime.Now.ToString("dd.MM.yyyy");
                Excel.Range range = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, 40]];
                range.Merge(Type.Missing);
                sheet.Cells[1, 1] = text_data;

                int row = 3;

                int year = Convert.ToInt32(cb_kv_year_start.Text);
                int kv = Convert.ToInt32(cb_kv_start.Text);
                if (have_code_err == false)
                {
                    if (checkBox2.Checked == true)
                    {
                        //по кварталам
                        sheet.Cells[2, 1] = "ПЕРЕОД";
                        sheet.Cells[2, 2] = "ОБЩЕЕ КОЛ-ВО, ШТ";
                        sheet.Cells[2, 3] = "КОЛ-ВО ПРИНЯТЫХ НА ГАРАНТИЙНЫЙ РЕМОНТ, ШТ";
                        sheet.Cells[2, 4] = "КОЛ-ВО НЕ ПРИНЯТЫХ НА ГАРАНТИЙНЫЙ РЕМОНТ, ШТ";
                        sheet.Cells[2, 5] = "ПРИНЯТЫЕ НА ГАРАНТИЙНЫЙ РЕМОНТ. ПРОЦЕНТ ОТ ВЫПУСКА";
                        sheet.Cells[2, 6] = "НЕ ПРИНЯТЫЕ НА ГАРАНТИЙНЫЙ РЕМОНТ. ПРОЦЕНТ ОТ ВЫПУСКА";
                        sheet.Cells[2, 7] = "ВСЕГО ВЫПУЩЕНО";

                        sheet.Columns.AutoFit();
                        sheet.Columns["A:A"].ColumnWidth = 14;

                        while (year <= Convert.ToInt32(cb_kv_year_end.Text))
                        {
                            sheet.Cells[row, 1] = rimsk[kv] + " кв. " + year + " г.";
                            row++;
                            kv++;
                            if (kv - 1 == Convert.ToInt32(cb_kv_end.Text) && year == Convert.ToInt32(cb_kv_year_end.Text))
                                kv = 5;
                            if (kv == 5)
                            {
                                sheet.Cells[row, 1] = "Всего за " + year + " г.";
                                range = sheet.Range[sheet.Cells[row, 1], sheet.Cells[row, 6]];
                                range.Font.Bold = true;
                                range.Interior.Color = Color.Gray;
                                row++;
                                kv = 1;
                                year++;
                            }
                        }

                        int last_row = 3;

                        for (int new_row = 3; new_row < row; new_row++)
                        {
                            for (int column = 2; column <= 6; column++)
                            {
                                if (column == 2)
                                {
                                    range = sheet.Cells[new_row, column] as Excel.Range;
                                    range.Formula = string.Format("=C" + new_row + "+D" + new_row);
                                }
                                else
                                    sheet.Cells[new_row, column] = 0;
                            }

                            range = sheet.Cells[new_row, 5] as Excel.Range;
                            range.Formula = string.Format("=C" + new_row + "/G" + new_row);
                            range = sheet.Cells[new_row, 6] as Excel.Range;
                            range.Formula = string.Format("=D" + new_row + "/G" + new_row);
                            range = sheet.Cells[new_row, 7] as Excel.Range;
                            range.Formula = string.Format("=B" + new_row);

                            range = sheet.Cells[new_row, 1] as Excel.Range;
                            string data_row = range.Value2;

                            if (data_row.IndexOf("Всего за") != -1)
                            {
                                range = sheet.Cells[new_row, 2] as Excel.Range;
                                range.Formula = string.Format("=SUM(B" + last_row + ":B" + (new_row - 1));

                                range = sheet.Cells[new_row, 3] as Excel.Range;
                                range.Formula = string.Format("=SUM(C" + last_row + ":C" + (new_row - 1));

                                range = sheet.Cells[new_row, 4] as Excel.Range;
                                range.Formula = string.Format("=SUM(D" + last_row + ":D" + (new_row - 1));

                                range = sheet.Cells[new_row, 5] as Excel.Range;
                                range.Formula = string.Format("=C" + new_row + "/G" + new_row);

                                range = sheet.Cells[new_row, 6] as Excel.Range;
                                range.Formula = string.Format("=D" + new_row + "/G" + new_row);

                                range = sheet.Cells[new_row, 7] as Excel.Range;
                                range.Formula = string.Format("=SUM(G" + last_row + ":G" + (new_row - 1));

                                last_row = new_row + 1;
                            }
                        }


                        while (data.Read())
                        {
                            DateTime dt = DateTime.Parse(data["DatePriem"].ToString());

                            kv = dt.Month < 4 ? 1 : dt.Month < 7 ? 2 : dt.Month < 10 ? 3 : 4;

                            for (int new_row = 3; new_row <= row; new_row++)
                            {
                                range = sheet.Cells[new_row, 1] as Excel.Range;
                                string data_row = range.Value2;

                                if (new_row != row && data_row.IndexOf(dt.Year.ToString()) != -1 && data_row.IndexOf(rimsk[kv]) != -1)
                                {
                                    int stolb = (data["Solution"].ToString().Trim() == "Гарантийный ремонт") ? 3 : 4;

                                    range = sheet.Cells[new_row, stolb] as Excel.Range;
                                    double a = range.Value2;
                                    a += 1;
                                    sheet.Cells[new_row, stolb] = a;
                                    break;
                                }
                            }
                        }
                    }
                    else
                    {
                        // создание шаблона по произвольной выборке даты
                    }
                }
                else
                {
                    /*
                     * Данный шаблон учитывает в расчётах неуникальные записи,
                     * то есть один счётчик может числиться сразу в ряде строк,
                     * поскольку каждый счётчик может иметь более одного кода ошибки/возврата.
                     */

                    if (checkBox2.Checked == true)
                    {
                        int years = Convert.ToInt32(cb_kv_year_end.Text) - Convert.ToInt32(cb_kv_year_start.Text);
                        int last_column = 0;
                        range = sheet.Range[sheet.Cells[2, 1], sheet.Cells[4, 1]];
                        range.Merge();
                        sheet.Cells[2, 1] = "№ п/п";

                        range = sheet.Range[sheet.Cells[2, 2], sheet.Cells[4, 2]];
                        range.Merge();
                        sheet.Cells[2, 2] = "Дефект";


                        range = sheet.Range[sheet.Cells[2, 3], sheet.Cells[2, 3 + years + 4]];
                        range.Merge();
                        sheet.Cells[2, 3] = "Процент к общему количеству приборов";

                        int start_year = Convert.ToInt32(cb_kv_year_start.Text);
                        int end_year = Convert.ToInt32(cb_kv_year_end.Text);

                        for (int i = start_year; i <= end_year; i++)
                        {
                            if (i != DateTime.Now.Year)
                            {
                                range = sheet.Range[sheet.Cells[3, 3 + i - start_year], sheet.Cells[4, 3 + i - start_year]];
                                range.Merge();
                                sheet.Cells[3, 3 + i - start_year] = i + " г.";
                            }
                            else
                            {
                                range = sheet.Range[sheet.Cells[3, 3 + i - start_year], sheet.Cells[3, 3 + i - start_year + 4]];
                                range.Merge();
                                sheet.Cells[3, 3 + i - start_year] = i + " г.";
                                sheet.Cells[4, 3 + i - start_year + 0] = "I кв.";
                                sheet.Cells[4, 3 + i - start_year + 1] = "II кв.";
                                sheet.Cells[4, 3 + i - start_year + 2] = "III кв.";
                                sheet.Cells[4, 3 + i - start_year + 3] = "IV кв.";
                                sheet.Cells[4, 3 + i - start_year + 4] = "Всего";

                                sheet.Cells[4, 3 + i - start_year + 4].Interior.Color =  Color.Gray;
                                sheet.Cells[4, 3 + i - start_year + 4].Font.Bold = true;
                                sheet.Cells[4, 3 + i - start_year + 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                sheet.Cells[4, 3 + i - start_year + 4].Borders.Color = ColorTranslator.ToOle(Color.Black);

                                last_column = 3 + i - start_year + 4;
                            }
                        }

                        for (row = 5; row < error_index + 5; row++)
                        {
                            sheet.Cells[row, 1] = row - 4;
                            sheet.Cells[row, 2] = errors[row - 5];

                            for (int col = 3; col <= last_column; col++)
                            {
                                sheet.Cells[row, col] = 0;

                                if (col == last_column)
                                {
                                    range = sheet.Cells[row, 11] as Excel.Range;
                                    range.Formula = string.Format("=SUM(C" + row + ":J" + (row - 1));

                                    sheet.Cells[row, col].Interior.Color = Color.Gray;
                                    sheet.Cells[row, col].Font.Bold = true;
                                    sheet.Cells[row, col].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                    sheet.Cells[row, col].Borders.Color = ColorTranslator.ToOle(Color.Black);
                                }
                                }
                        }
                        range = sheet.Cells[1, 2] as Excel.Range;
                        range.EntireColumn.AutoFit();

                        while (data.Read()) {

                            DateTime dt = DateTime.Parse(data["DatePriem"].ToString());

                            kv = dt.Month < 4 ? 1 : dt.Month < 7 ? 2 : dt.Month < 10 ? 3 : 4;

                            for (int new_row = 5; new_row <= row; new_row++) {

                                range = sheet.Cells[new_row, 2] as Excel.Range;
                                string dataCodeB = Convert.ToString(range.Value2);

                                //var dataCodeB = (sheet.Cells[new_row, 2] as Excel.Range)?.Value2?.ToString() ?? "";

                                if (new_row != row && data["Codeb"].ToString().IndexOf(dataCodeB) != -1) {

                                    for (int index = 0, column = 3; start_year <= end_year; start_year++, index++) {

                                        if (dt.Year == start_year) {

                                            column += index;

                                            if (dt.Year == end_year) column = dt.Month < 4 ? 7: dt.Month < 7 ? 8 : dt.Month < 10 ? 9 : 10;

                                            range = sheet.Cells[new_row, column] as Excel.Range;
                                            double a = range.Value2;
                                            a += 1;
                                            sheet.Cells[new_row, column] = a;

                                            break;
                                        }
                                    }
                                    break;
                                }
                            }
                        }
                    }
                    else { }
                }
            }
            comm.Connection.Close();

        }

        private void dataGridView2_RowEnter_1(object sender, DataGridViewCellEventArgs e)
        {
            cust_id.Text = dataGridView2["Column1", e.RowIndex].Value.ToString();
        }

        void Form_acts(string act_IDParty)
        {
            MySqlDataReader max_id = ExecutQuery_Select("SELECT MAX(ID) as Max_ID FROM acts");
            int act_ID = 0;
            string act_Date = "";
            string act_sposob_dost = "";
            string act_num_dost = "";
            string act_date_dost = "";
            string act_customer = "";
            //string act_meters = "";

            //meter[] meters_;


            customer cust = null;
            string customerid = "";

            if (max_id != null)
            {
                while (max_id.Read())
                {
                    if (max_id["Max_ID"].ToString().Trim() != "")
                    {
                        act_ID = Convert.ToInt32(max_id["Max_ID"].ToString());
                    }
                }
            }
            act_ID++;
            comm.Connection.Close();
            MySqlDataReader find_date = ExecutQuery_Select("SELECT * FROM inwork WHERE IDParty='" + act_IDParty + "' ORDER BY DatePriem DESC");
            if (find_date != null)
            {
                while (find_date.Read())
                {
                    act_Date = find_date["DatePriem"].ToString();
                    act_sposob_dost = find_date["sposob_dost"].ToString();
                    act_num_dost = find_date["num_dost"].ToString();
                    act_date_dost = find_date["date_dost"].ToString();
                    customerid = find_date["CustomerID"].ToString();
                    break;
                }
            }
            comm.Connection.Close();


            MySqlDataReader find_customer = ExecutQuery_Select("SELECT * FROM customers WHERE Descr = '" + customerid + "'");
            if (find_customer != null)
            {
                while (find_customer.Read())
                {
                    cust = new customer
                    {
                        ID = find_customer["ID"].ToString(),
                        Descr = find_customer["Descr"].ToString(),
                        ContFace = find_customer["ContFace"].ToString(),
                        Phone = find_customer["Phone"].ToString(),
                        _Index = find_customer["_Index"].ToString(),
                        Resp = find_customer["Resp"].ToString(),
                        Oblast = find_customer["Oblast"].ToString(),
                        City = find_customer["City"].ToString(),
                        Street = find_customer["Street"].ToString(),
                        Num_h = find_customer["Num_h"].ToString(),
                        Num_f = find_customer["Num_f"].ToString()
                    };
                    break;
                }
            }
            comm.Connection.Close();
            //act_customer = JsonSerializer.Serialize<customer>(cust);
            act_customer = "{" +
                "\"ID\":" + "\"" + cust.ID + "\"," +
                "\"Descr\":" + "\"" + cust.Descr + "\"," +
                "\"ContFace\":" + "\"" + cust.ContFace + "\"," +
                "\"Phone\":" + "\"" + cust.Phone + "\"," +
                "\"_Index\":" + "\"" + cust._Index + "\"," +
                "\"Resp\":" + "\"" + cust.Resp + "\"," +
                "\"Oblast\":" + "\"" + cust.Oblast + "\"," +
                "\"City\":" + "\"" + cust.City + "\"," +
                "\"Street\":" + "\"" + cust.Street + "\"," +
                "\"Num_h\":" + "\"" + cust.Num_h + "\"," +
                "\"Num_f\":" + "\"" + cust.Num_f + "\"" +
                "}";



            List<string> reasons = new List<string>();
            List<string> like_types_meters = new List<string>();
            like_types_meters.Add("(TypeMeter LIKE '%СХВ%' OR TypeMeter LIKE '%СГВ%')");
            like_types_meters.Add("(TypeMeter LIKE '%СГБМ%' OR TypeMeter LIKE '%СГБУ%')");
            like_types_meters.Add("(TypeMeter LIKE '%СВМ%')");
            like_types_meters.Add("(TypeMeter LIKE '%ЭСО%' OR TypeMeter LIKE '%БАРС%')");
            like_types_meters.Add("(TypeMeter LIKE '%СТК%')");
            like_types_meters.Add("(TypeMeter LIKE '%ПВМ%')");
            like_types_meters.Add("(TypeMeter LIKE '%РД%')");


            MySqlDataReader find_reasons = ExecutQuery_Select("SELECT * FROM reasonreturn");
            if (find_reasons != null)
            {
                while (find_reasons.Read())
                {
                    reasons.Add(find_reasons["Reason"].ToString().Trim());
                }
            }
            comm.Connection.Close();
            List<meter> meters = new List<meter>();
            foreach (string type_meter in like_types_meters)
            {
                foreach (string reason in reasons)
                {
                    string type_prib = type_meter;

                    if (type_prib.IndexOf("СХВ") != -1 || type_prib.IndexOf("СГВ") != -1) type_prib = "СХВ/СГВ";
                    else if (type_prib.IndexOf("СГБМ") != -1 || type_prib.IndexOf("СГБУ") != -1) type_prib = "СГБМ/СГБУ";
                    else if (type_prib.IndexOf("СВМ") != -1) type_prib = "СВМ";
                    else if (type_prib.IndexOf("ЭСО") != -1 || type_prib.IndexOf("БАРС") != -1) type_prib = "ЭСО/БАРС";
                    else if (type_prib.IndexOf("СТК") != -1) type_prib = "СТК";
                    else if (type_prib.IndexOf("ПВМ") != -1) type_prib = "ПВМ";
                    else type_prib = "РД";
                    meters.Clear();
                    MySqlDataReader find_meter = ExecutQuery_Select("SELECT * FROM inwork WHERE IDParty='" + act_IDParty + "' AND " + type_meter + " AND Solution = '" + reason + "'");
                    if (find_meter != null)
                    {
                        while (find_meter.Read())
                        {

                            meters.Add(new meter()
                            {
                                IDParty = act_IDParty,
                                Ser_Num = find_meter["Ser_Num"].ToString(),
                                TypeMeter = find_meter["TypeMeter"].ToString(),
                                DateCreate = find_meter["DateCreate"].ToString(),
                                Solution = find_meter["Solution"].ToString(),
                                Codeb = find_meter["Codeb"].ToString(),
                                CustomerID = find_meter["CustomerID"].ToString(),
                                UserID = find_meter["UserID"].ToString(),
                                DatePriem = find_meter["DatePriem"].ToString(),
                                DateAnaliz = find_meter["DateAnaliz"].ToString(),
                                Descr = find_meter["Descr"].ToString(),
                                kit = find_meter["kit"].ToString(),
                                sposob_dost = find_meter["sposob_dost"].ToString(),
                                num_dost = find_meter["num_dost"].ToString(),
                                date_dost = find_meter["date_dost"].ToString(),
                                narabotka = find_meter["narabotka"].ToString(),
                                Code_err_user = find_meter["Code_err_user"].ToString()
                            });
                            //MessageBox.Show(find_meter["ID"] + " " + find_meter["IDParty"] + " " + find_meter["TypeMeter"] + " " + find_meter["Solution"] + " ");
                        }
                        comm.Connection.Close();
                        if (meters.Count < 1) continue;
                        string data_meter_json = "[";
                        for (int i = 0; i < meters.Count; i++)
                        {
                            data_meter_json += "{" +
                                "\"IDParty\":" + "\"" + meters[i].IDParty + "\"," +
                                "\"Ser_Num\":" + "\"" + meters[i].Ser_Num + "\"," +
                                "\"TypeMeter\":" + "\"" + meters[i].TypeMeter + "\"," +
                                "\"DateCreate\":" + "\"" + meters[i].DateCreate + "\"," +
                                "\"Solution\":" + "\"" + meters[i].Solution + "\"," +
                                "\"Codeb\":" + "\"" + meters[i].Codeb + "\"," +
                                "\"CustomerID\":" + "\"" + meters[i].CustomerID + "\"," +
                                "\"UserID\":" + "\"" + meters[i].UserID + "\"," +
                                "\"DatePriem\":" + "\"" + meters[i].DatePriem + "\"," +
                                "\"DateAnaliz\":" + "\"" + meters[i].DateAnaliz + "\"," +
                                "\"Descr\":" + "\"" + meters[i].Descr + "\"," +
                                "\"kit\":" + "\"" + meters[i].kit + "\"," +
                                "\"sposob_dost\":" + "\"" + meters[i].sposob_dost + "\"," +
                                "\"num_dost\":" + "\"" + meters[i].num_dost + "\"," +
                                "\"date_dost\":" + "\"" + meters[i].date_dost + "\"," +
                                "\"narabotka\":" + "\"" + meters[i].narabotka + "\"," +
                                "\"Code_err_user\":" + "\"" + meters[i].Code_err_user + "\"" +
                                "}";
                            if (i != meters.Count - 1) data_meter_json += ",";
                        }
                        data_meter_json += "]";
                        ExecutQuery_Insert("INSERT INTO `acts`" +
                            "(`ID`, `Date`, `IDParty`," +
                            " `sposob_dost`, `num_dost`, `date_dost`," +
                            " `ID_Customer`, `Schetciki`, `Solution`, `TypeMeters`, `ID_User`)" +
                            " VALUES ('" + act_ID + "','" + meters[0].DatePriem + "','" + act_IDParty + "'," +
                            "'" + act_sposob_dost + "','" + act_num_dost + "','" + act_date_dost + "'," +
                            "'" + act_customer + "','" + data_meter_json + "','" + reason + "','" + type_prib + "','" + meters[0].UserID + "')");
                        act_ID++;
                    }
                    comm.Connection.Close();
                    //MessageBox.Show("Новое");
                }
            }
            /*
            int count_meters=0;
            MySqlDataReader find_meters = ExecutQuery_Select("SELECT COUNT(*) as count FROM inwork WHERE IDParty = '" + act_IDParty + "'");
            if (find_meters != null)
            {
                while (find_meters.Read())
                {
                    count_meters = Convert.ToInt32(find_meters["count"].ToString());
                    break;
                }
            }
            comm.Connection.Close();
            meters_ = new meter[count_meters];
            int i = 0;
            find_meters = ExecutQuery_Select("SELECT * FROM inwork WHERE IDParty = '" + act_IDParty + "'");
            if (find_meters != null)
            {
                while (find_meters.Read())
                {
                    meters_[i] = new meter
                    {
                        IDParty = act_IDParty,
                        Ser_Num=find_meters["Ser_Num"].ToString(),
                        TypeMeter = find_meters["TypeMeter"].ToString(),
                        DateCreate = find_meters["DateCreate"].ToString(),
                        Solution = find_meters["Solution"].ToString(),
                        Codeb = find_meters["Codeb"].ToString(),
                        CustomerID = find_meters["CustomerID"].ToString(),
                        UserID = find_meters["UserID"].ToString(),
                        DatePriem = find_meters["DatePriem"].ToString(),
                        DateAnaliz = find_meters["DateAnaliz"].ToString(),
                        Descr = find_meters["Descr"].ToString(),
                        kit = find_meters["kit"].ToString(),
                        sposob_dost = find_meters["sposob_dost"].ToString(),
                        num_dost = find_meters["num_dost"].ToString(),
                        date_dost = find_meters["date_dost"].ToString()
                    };
                    i++;
                }
            }
            comm.Connection.Close();



            
            act_meters = JsonSerializer.Serialize<meter[]>(meters_);
            act_ID++;
            */
        }
        void Show_All_Acts(int IDParty)
        {
            List<int> id_acts_int = new List<int>();
            MySqlDataReader id_acts = ExecutQuery_Select("SELECT * FROM acts WHERE IDParty = '" + IDParty + "'");
            if (id_acts != null)
            {
                while (id_acts.Read())
                {
                    id_acts_int.Add(Convert.ToInt32(id_acts["ID"].ToString().Trim()));
                }
            }
            comm.Connection.Close();
            for (int i = 0; i < id_acts_int.Count; i++)
            {
                Show_Act(id_acts_int[i]);
            }
        }
        void Show_Act(int ID_Act)
        {

            MySqlDataReader reader1 = ExecutQuery_Select("SELECT * FROM acts WHERE ID = '" + ID_Act + "'");
            if (reader1 != null)
            {
                if (!reader1.HasRows)
                {
                    MessageBox.Show("Акт не найден.");
                    comm.Connection.Close();
                    return;
                }
            }
            comm.Connection.Close();

            Word.Application appWord;
            Word.Document docWord = null;
            object missobj = System.Reflection.Missing.Value;
            object falseobj = false;
            object trueobj = true;

            appWord = new Word.Application();
            object path_sh = StartPath + "Shablon_Act_priemki.docx";
            try
            {
                docWord = appWord.Documents.Add(ref path_sh, ref missobj, ref missobj, ref missobj);
            }
            catch (Exception err)
            {
                docWord.Close(ref falseobj, ref missobj, ref missobj);
                appWord.Quit(ref missobj, ref missobj, ref missobj);
                docWord = null;
                appWord = null;
                throw err;
            }
            appWord.Visible = true;

            object date_act = "date_act";
            object num_act = "num_act";
            object num_pn_act = "num_pn_act";
            object date_act2 = "date_act2";
            object sp_pol_act = "sp_pol_act";
            object sp_pol_num_act = "sp_pol_num_act";
            object sp_pol_date_act = "sp_pol_date_act";

            object customer_act = "customer_act";
            object cust_index_act = "cust_index_act";
            object cust_resp_act = "cust_resp_act";
            object cust_oblast_act = "cust_oblast_act";
            object cust_city_act = "cust_city_act";
            object cust_street_act = "cust_street_act";
            object cust_h_act = "cust_h_act";
            object cust_flat_act = "cust_flat_act";

            object ser_prib1_act = "ser_prib1_act";
            object ser_prib2_act = "ser_prib2_act";
            object ser_prib3_act = "ser_prib3_act";
            object ser_prib4_act = "ser_prib4_act";
            object ser_prib5_act = "ser_prib5_act";

            object negarant_act = "negarant_act";
            object user_act = "user_act";
            object date_act3 = "date_act3";
            object solution_act = "solution_act";

            Word.Bookmark bookmark_date_act = docWord.Bookmarks[ref date_act];
            Word.Bookmark bookmark_num_act = docWord.Bookmarks[ref num_act];
            Word.Bookmark bookmark_num_pn_act = docWord.Bookmarks[ref num_pn_act];
            Word.Bookmark bookmark_date_act2 = docWord.Bookmarks[ref date_act2];
            Word.Bookmark bookmark_sp_pol_act = docWord.Bookmarks[ref sp_pol_act];
            Word.Bookmark bookmark_sp_pol_num_act = docWord.Bookmarks[ref sp_pol_num_act];
            Word.Bookmark bookmark_sp_pol_date_act = docWord.Bookmarks[ref sp_pol_date_act];

            Word.Bookmark bookmark_customer_act = docWord.Bookmarks[ref customer_act];
            Word.Bookmark bookmark_cust_index_act = docWord.Bookmarks[ref cust_index_act];
            Word.Bookmark bookmark_cust_resp_act = docWord.Bookmarks[ref cust_resp_act];
            Word.Bookmark bookmark_cust_oblast_act = docWord.Bookmarks[ref cust_oblast_act];
            Word.Bookmark bookmark_cust_city_act = docWord.Bookmarks[ref cust_city_act];
            Word.Bookmark bookmark_cust_street_act = docWord.Bookmarks[ref cust_street_act];
            Word.Bookmark bookmark_cust_h_act = docWord.Bookmarks[ref cust_h_act];
            Word.Bookmark bookmark_cust_flat_act = docWord.Bookmarks[ref cust_flat_act];

            Word.Bookmark bookmark_ser_prib1_act = docWord.Bookmarks[ref ser_prib1_act];
            Word.Bookmark bookmark_ser_prib2_act = docWord.Bookmarks[ref ser_prib2_act];
            Word.Bookmark bookmark_ser_prib3_act = docWord.Bookmarks[ref ser_prib3_act];
            Word.Bookmark bookmark_ser_prib4_act = docWord.Bookmarks[ref ser_prib4_act];
            Word.Bookmark bookmark_ser_prib5_act = docWord.Bookmarks[ref ser_prib5_act];

            Word.Bookmark bookmark_negarant_act = docWord.Bookmarks[ref negarant_act];
            Word.Bookmark bookmark_user_act = docWord.Bookmarks[ref user_act];
            Word.Bookmark bookmark_date_act3 = docWord.Bookmarks[ref date_act3];
            Word.Bookmark bookmark_solution_act = docWord.Bookmarks[ref solution_act];

            string bd_ID_act = "";
            string bd_date_act = "";
            string bd_IDParty_act = "";
            string bd_sp_dost_act = "";
            string bd_num_dost_act = "";
            string bd_date_dost_act = "";
            string bd_ID_Customer_JSON = "";
            string bd_ID_Schetciki_JSON = "";
            string bd_Solution_act = "";
            string bd_ID_User_act = "";


            MySqlDataReader readact = ExecutQuery_Select("SELECT * FROM acts WHERE ID = '" + ID_Act + "'"); ;
            if (readact == null) return;
            if (readact.HasRows != false)
            {
                while (readact.Read())
                {
                    bd_ID_act = readact["ID"].ToString();
                    bd_date_act = readact["Date"].ToString();
                    bd_IDParty_act = readact["IDParty"].ToString();
                    bd_sp_dost_act = readact["sposob_dost"].ToString();
                    bd_num_dost_act = readact["num_dost"].ToString();
                    bd_date_dost_act = readact["date_dost"].ToString();
                    bd_ID_Customer_JSON = readact["ID_Customer"].ToString();
                    bd_ID_Schetciki_JSON = readact["Schetciki"].ToString();
                    bd_Solution_act = readact["Solution"].ToString();
                    bd_ID_User_act = readact["ID_User"].ToString();
                }
            }
            comm.Connection.Close();

            bookmark_date_act.Range.Text = bookmark_date_act2.Range.Text = bookmark_date_act3.Range.Text = bd_date_act.Substring(0, 10);
            bookmark_num_act.Range.Text = bd_ID_act;
            bookmark_num_pn_act.Range.Text = bd_IDParty_act;
            bookmark_sp_pol_act.Range.Text = bd_sp_dost_act;
            bookmark_sp_pol_num_act.Range.Text = bd_num_dost_act;
            bookmark_sp_pol_date_act.Range.Text = bd_date_dost_act;


            //meter meter1 = new meter() { Descr = "Лалка" };
            //string json = JsonConvert.SerializeObject(meter1);

            //MessageBox.Show(json);
            //meter met2 = JsonConvert.DeserializeObject<meter>(json);
            //MessageBox.Show(met2.Descr);

            MessageBox.Show(bd_ID_Customer_JSON);
            customer cust = JsonConvert.DeserializeObject<customer>(bd_ID_Customer_JSON);

            bookmark_customer_act.Range.Text = cust.Descr;
            bookmark_cust_index_act.Range.Text = cust._Index;
            bookmark_cust_resp_act.Range.Text = cust.Resp;
            bookmark_cust_oblast_act.Range.Text = cust.Oblast;
            bookmark_cust_city_act.Range.Text = cust.City;
            bookmark_cust_street_act.Range.Text = cust.Street;
            bookmark_cust_h_act.Range.Text = cust.Num_h;
            bookmark_cust_flat_act.Range.Text = cust.Num_f;

            if (bd_Solution_act == "Негарантия") bookmark_negarant_act.Range.Text = "V";

            MySqlDataReader readusr = ExecutQuery_Select("SELECT * FROM users WHERE ID = '" + bd_ID_User_act + "'"); ;
            if (readusr == null) return;
            if (readusr.HasRows != false)
            {
                while (readusr.Read())
                {
                    bookmark_user_act.Range.Text = readusr["position"].ToString() + " " + readusr["Descr"].ToString();
                }
            }
            comm.Connection.Close();
            bookmark_solution_act.Range.Text = bd_Solution_act;
            int all_prib_without_kit = 0;
            int all_prib_with_kit = 0;
            meter[] meters = JsonConvert.DeserializeObject<meter[]>(bd_ID_Schetciki_JSON);
            int row = 2;
            int column = 3;
            for (int i = 0; i < meters.Length; i++)
            {
                int sum_all = 0;
                int sum_kompl = 0;
                string type_meter_cikle = meters[i].TypeMeter;
                if (type_meter_cikle == "none") continue;
                for (int a = 0; a < meters.Length; a++)
                {
                    if (meters[a].TypeMeter == type_meter_cikle)
                    {
                        sum_all++;
                        all_prib_without_kit++;
                        if (meters[a].kit.Trim() == "1")
                        {
                            sum_kompl++;
                            all_prib_with_kit++;
                        }
                        meters[a].TypeMeter = "none";
                    }
                }
                if (row >= 4) column = 4;
                if (row >= 10) column = 6;
                docWord.Tables[2].Cell(row, column).Range.Text = type_meter_cikle;
                docWord.Tables[2].Cell(row, column + 1).Range.Text = sum_all.ToString();
                docWord.Tables[2].Cell(row, column + 2).Range.Text = sum_kompl.ToString();
                row++;
                if (row >= 10)
                    docWord.Tables[2].Rows.Add(ref missobj);
                if (i == meters.Length - 1)
                {
                    if (i < 10)
                    {
                        row = 10;
                        column = 6;
                    }
                    docWord.Tables[2].Cell(row, column).Range.Text = "Всего:";
                    docWord.Tables[2].Cell(row, column + 1).Range.Text = all_prib_without_kit.ToString();
                    docWord.Tables[2].Cell(row, column + 2).Range.Text = all_prib_with_kit.ToString();
                }
            }

            Word.Bookmark[] bms = new Word.Bookmark[] { bookmark_ser_prib1_act, bookmark_ser_prib2_act, bookmark_ser_prib3_act, bookmark_ser_prib4_act, bookmark_ser_prib5_act };

            if (all_prib_without_kit <= 5)
            {
                for (int i = 0; i < meters.Length; i++)
                {
                    bms[i].Range.Text = meters[i].Ser_Num;
                }
            }

            //docWord.Tables[2].Cell(2, 3).Range.Text = "Test";
            /*
            

            Word.Bookmark bookmark_ser_prib1_act = docWord.Bookmarks[ref ser_prib1_act];
            Word.Bookmark bookmark_ser_prib2_act = docWord.Bookmarks[ref ser_prib2_act];
            Word.Bookmark bookmark_ser_prib3_act = docWord.Bookmarks[ref ser_prib3_act];
            Word.Bookmark bookmark_ser_prib4_act = docWord.Bookmarks[ref ser_prib4_act];
            Word.Bookmark bookmark_ser_prib5_act = docWord.Bookmarks[ref ser_prib5_act];

            Word.Bookmark bookmark_negarant_act = docWord.Bookmarks[ref negarant_act];
            Word.Bookmark bookmark_user_act = docWord.Bookmarks[ref user_act];
            Word.Bookmark bookmark_date_act3 = docWord.Bookmarks[ref date_act3];
            Word.Bookmark bookmark_solution_act = docWord.Bookmarks[ref solution_act];*/



        }

        private void act_priem_num_PN_TextChanged(object sender, EventArgs e)
        {

        }

        private void act_priem_num_acta_Click(object sender, EventArgs e)
        {
            act_priem_num_acta.Items.Clear();
            if (act_priem_num_PN.Text.Trim() == "") return;

            MySqlDataReader find_acts = ExecutQuery_Select("SELECT * FROM acts WHERE IDParty = '" + act_priem_num_PN.Text + "'"); ;
            if (find_acts == null) return;
            if (find_acts.HasRows != false)
            {
                while (find_acts.Read())
                {
                    act_priem_num_acta.Items.Add(find_acts["ID"]);
                }
            }
            comm.Connection.Close();

        }

        private void Find_Act_Click(object sender, EventArgs e)
        {
            Show_Act(Convert.ToInt32(act_priem_num_acta.Text));
        }

        private void cb_coded_TextChanged(object sender, EventArgs e)
        {
            /*
            if (cb_coded.Text.Trim() == "")
            {
                cb_coded.Items.AddRange(code_err_water);
            }
            else
            {
                foreach (string code_err in code_err_water)
                {
                    if (code_err.Trim() == cb_coded.Text.Trim())
                    {
                        cb_coded.Items.AddRange(code_err_water);
                        return;
                    }
                }
                //MessageBox.Show("Clear data = "+cb_coded.Text);
                cb_coded.Items.Clear();
                foreach (string code_err in code_err_water)
                {
                    if (code_err.IndexOf(cb_coded.Text) != -1)
                    {
                        cb_coded.Items.Add(code_err);
                        cb_coded.DroppedDown = true;
                        //cb_coded.Focus();
                    }
                }
            }
            cb_coded.SelectionStart = cb_coded.Text.Length;*/
        }

        private void cb_coded_Click(object sender, EventArgs e)
        {

            if (cb_coded.Text.Trim() == "")
            {
                cb_coded.Items.Clear();
                cb_coded.Items.AddRange(code_err_water);
            }
        }

        private void cb_coded_SelectedIndexChanged(object sender, EventArgs e)
        {
            ActiveControl = null;
            //cb_coded.SelectionStart = 0;
        }

        private void cb_coded_SelectionChangeCommitted(object sender, EventArgs e)
        {
        }

        private void label11_Click(object sender, EventArgs e)
        {
        }

        private void tb_CLA_num_act_TextChanged(object sender, EventArgs e)
        {
            string num_act = tb_CLA_num_act.Text;
            label30.Text = label32.Text = "";
            MySqlDataReader act = ExecutQuery_Select("SELECT * FROM acts WHERE ID = '" + num_act + "'");
            if (act != null)
            {
                while (act.Read())
                {
                    label30.Text = act["TypeMeters"] + " " + act["Solution"];
                    label32.Text = "№ партии - " + act["IDParty"].ToString();
                }
            }
            comm.Connection.Close();


        }

        private void btn_find_CLA_Click(object sender, EventArgs e)
        {
            string num_act = tb_CLA_num_act.Text;
            MySqlDataReader reader1 = ExecutQuery_Select("SELECT * FROM acts WHERE ID = '" + num_act + "'");
            if (reader1 != null)
            {
                if (!reader1.HasRows)
                {
                    MessageBox.Show("Акт, для формирования контрольного листка анализа не найден.");
                    comm.Connection.Close();
                    return;
                }
            }
            comm.Connection.Close();

            if (label30.Text.IndexOf("СХВ/СГВ") != -1) Contr_list_analiza_SHV_SGV(num_act);
            else MessageBox.Show("Пока готовы акты по СВ");
        }

        void Contr_list_analiza_SHV_SGV(string num_act)
        {

            Word.Application appWord;
            Word.Document docWord = null;
            object missobj = System.Reflection.Missing.Value;
            object falseobj = false;
            object trueobj = true;

            appWord = new Word.Application();
            object path_sh = StartPath + "Kontr_list_analiza_shablon_SHV_SGV.docx";
            try
            {
                docWord = appWord.Documents.Add(ref path_sh, ref missobj, ref missobj, ref missobj);
            }
            catch (Exception err)
            {
                docWord.Close(ref falseobj, ref missobj, ref missobj);
                appWord.Quit(ref missobj, ref missobj, ref missobj);
                docWord = null;
                appWord = null;
                throw err;
            }
            appWord.Visible = true;
            object ref_num = "num_act";
            object ref_kolvo = "kol_vo";
            object ref_customer = "customer";
            object ref_date_an = "Date_analiz";
            Word.Bookmark bookmark_ref_num = docWord.Bookmarks[ref ref_num];
            Word.Bookmark bookmark_ref_kolvo = docWord.Bookmarks[ref ref_kolvo];
            Word.Bookmark bookmark_ref_customer = docWord.Bookmarks[ref ref_customer];
            Word.Bookmark bookmark_ref_date_an = docWord.Bookmarks[ref ref_date_an];


            Word.Table tbl = docWord.Tables[2];
            int row = 2;
            int kol_vo = 0;

            string bd_schitciki_JSON = "";
            string bd_customer_JSON = "";

            MySqlDataReader act = ExecutQuery_Select("SELECT * FROM acts WHERE ID = '" + num_act + "'");
            if (act != null)
            {
                while (act.Read())
                {
                    bd_schitciki_JSON = act["Schetciki"].ToString();
                    bd_customer_JSON = act["ID_Customer"].ToString();
                }
            }
            comm.Connection.Close();

            bookmark_ref_customer.Range.Text = JsonConvert.DeserializeObject<customer>(bd_customer_JSON).Descr;
            bookmark_ref_num.Range.Text = num_act;

            meter[] meters = JsonConvert.DeserializeObject<meter[]>(bd_schitciki_JSON);

            foreach (meter my_meter in meters)
            {
                if (row == 2) bookmark_ref_date_an.Range.Text = my_meter.DateAnaliz.Substring(0, 10);
                if (row == 16)
                {
                    tbl.Rows.Add(ref missobj);
                }
                tbl.Cell(row, 1).Range.Text = (row - 1).ToString();
                tbl.Cell(row, 2).Range.Text = my_meter.Code_err_user;
                tbl.Cell(row, 3).Range.Text = my_meter.TypeMeter;
                tbl.Cell(row, 4).Range.Text = my_meter.Ser_Num;
                tbl.Cell(row, 5).Range.Text = my_meter.DateCreate.Substring(0, 10);
                tbl.Cell(row, 6).Range.Text = my_meter.narabotka;
                tbl.Cell(row, 7).Range.Text = my_meter.Codeb.Substring(0, my_meter.Codeb.Length - 1).Replace('|', '\n');
                tbl.Cell(row, 8).Range.Text = my_meter.Descr;
                row++;
                kol_vo++;
            }
            bookmark_ref_kolvo.Range.Text = kol_vo.ToString();


        }

        private void label24_Click(object sender, EventArgs e)
        {

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void btn_showCD_Click(object sender, EventArgs e)
        {
            Chose_CD _CD = new Chose_CD(code_err_water, ref cb_coded);
            _CD.ShowDialog();
        }

        private void label23_Click(object sender, EventArgs e)
        {

        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            checkedListBox1.Items.Clear();

            if (tabControl1.SelectedIndex == 1)
            {
                MySqlDataReader reader1 = ExecutQuery_Select("SELECT * FROM `inwork` GROUP BY TypeMeter");
                if (reader1 != null)
                {
                    if (reader1.HasRows)
                    {
                        while (reader1.Read())
                        {
                            checkedListBox1.Items.Add(reader1["TypeMeter"].ToString());
                            //MessageBox.Show("add + " + reader1["TypeMeter"].ToString());
                        }
                    }
                }
                comm.Connection.Close();
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                cb_kv_end.Enabled = cb_kv_start.Enabled = cb_kv_year_end.Enabled = cb_kv_year_start.Enabled = true;
                textBox1.Enabled = textBox2.Enabled = monthCalendar2.Enabled = monthCalendar1.Enabled = false;
            }
            else
            {
                cb_kv_end.Enabled = cb_kv_start.Enabled = cb_kv_year_end.Enabled = cb_kv_year_start.Enabled = false;
                textBox1.Enabled = textBox2.Enabled = monthCalendar2.Enabled = monthCalendar1.Enabled = true;

            }
        }

        private void monthCalendar1_DateChanged(object sender, DateRangeEventArgs e)
        {
            textBox1.Text = e.Start.ToString("yyyy-MM-dd");
            monthCalendar2.MinDate = e.Start;
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            //monthCalendar1.Visible = false;

        }

        private void textBox1_Enter(object sender, EventArgs e)
        {
            // monthCalendar1.Visible = true;
        }

        private void textBox2_Enter(object sender, EventArgs e)
        {
            //monthCalendar2.Visible = true;

        }

        private void textBox2_Leave(object sender, EventArgs e)
        {
            //monthCalendar2.Visible = false;

        }

        private void monthCalendar2_DateChanged(object sender, DateRangeEventArgs e)
        {
            textBox2.Text = e.Start.ToString("yyyy-MM-dd");
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
    class act
    {
        public string ID { get; set; }
        public string Date { get; set; }
        public string IDParty { get; set; }
        public string sposob_dost { get; set; }
        public string num_dost { get; set; }
        public string date_dost { get; set; }
        public string ID_Customer { get; set; }
        public string Schetciki { get; set; }
        public string Solution { get; set; }
        public string ID_User { get; set; }
    }
    class meter
    {
        public string IDParty { get; set; }
        public string Ser_Num { get; set; }
        public string TypeMeter { get; set; }
        public string DateCreate { get; set; }
        public string Solution { get; set; }
        public string Codeb { get; set; }
        public string Code_err_user { get; set; }
        public string CustomerID { get; set; }
        public string UserID { get; set; }
        public string DatePriem { get; set; }
        public string DateAnaliz { get; set; }
        public string Descr { get; set; }
        public string kit { get; set; }
        public string sposob_dost { get; set; }
        public string num_dost { get; set; }
        public string date_dost { get; set; }
        public string narabotka { get; set; }
    }
    class customer
    {
        public string ID { get; set; }
        public string Descr { get; set; }
        public string ContFace { get; set; }
        public string Phone { get; set; }
        public string _Index { get; set; }
        public string Resp { get; set; }
        public string Oblast { get; set; }
        public string City { get; set; }
        public string Street { get; set; }
        public string Num_h { get; set; }
        public string Num_f { get; set; }
    }
}
