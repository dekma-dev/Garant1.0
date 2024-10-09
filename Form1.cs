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
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;

namespace Garant1._0
{
    public partial class Form1 : Form
    {
        public MySqlCommand comm;

        int ID_User;

        bool find_row = false;

        string StartPath;

        string[] code_err_water;

        public Form1(MySqlCommand comm)
        {
            InitializeComponent();code_err_water = new string[] {
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

            StartPath = Application.StartupPath + "\\";

            this.comm = comm;

            SetSQLMode();
            Find_users();

            defectCodesList.Items.Clear();
            foreach (string code in code_err_water)
            {
                defectCodesList.Items.Add(code);
            }

            label24.Visible = false;
            label25.Visible = false;
            mailNumberTB.Visible = false;
            sendDateTB.Visible = false;
            sendDateTB.Text = "";
            mailNumberTB.Text = "";

            Create_DataGridView();
            fill_cb();
            FillDataGridViewCustomer();
            reportYearEnd.Text = DateTime.Now.ToString("yyyy");
            dateFromTB.Enabled = dateToTB.Enabled = monthCalendarTo.Enabled = monthCalendarFrom.Enabled = false;
        }

        void Find_users()
        {
            MySqlDataReader reader = ExecutQuery_Select("SELECT * FROM users");
            if (reader != null)
                while (reader.Read())
                {
                    choseUserComboBox.Items.Add(reader["Descr"].ToString());
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
            MySqlDataReader reader = null;
            try
            {
                comm.CommandText = query;
                if (comm.Connection.State == ConnectionState.Closed)
                {
                    comm.Connection.Open();
                }
                reader = comm.ExecuteReader();
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                richTextBox1.Text += ("err query = " + query + "\nData\"" + ex.Message + "\"\n");
                MessageBox.Show(ex.Message);
            }
            return reader;
        }

        int ExecutQuery_Insert(String query)
        {
            int i = 0;
            try
            {
                comm.CommandText = query;
                comm.Connection.Open();
                i = comm.ExecuteNonQuery();
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
            if (choseUserComboBox.Text.Trim() != "") {
                tabControl1.Enabled = true;
                MySqlDataReader reader = ExecutQuery_Select("SELECT * FROM users WHERE Descr = '" + choseUserComboBox.Text + "'");
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

            chosePartysNumber.Items.Clear();
            defectCode4FindAct.Items.Clear();
            Refresh_DataGridView();

            chosePartysNumber.Items.Add("0 - Новая партия");
            MySqlDataReader reader2 = ExecutQuery_Select("SELECT DISTINCT IDParty FROM inwork");
            if (reader2 == null) return;
            if (reader2.HasRows != false) {

                while (reader2.Read()) {
                    chosePartysNumber.Items.Add(reader2["IDParty"].ToString());
                    defectCode4FindAct.Items.Add(reader2["IDParty"].ToString());
                }
            }
            comm.Connection.Close();

            Refresh_DataGridView();
        }

        private async void tb_Serial_num_TextChanged(object sender, EventArgs e)
        {
            string serial_num = serialNumberTB.Text.Trim();
            producedDateTB.Text = "";
            metersTypeTB.Text = "";
            defectCodeTB.Items.Clear();
            if (serial_num.Length == 8) {

                await Task.Delay(10);
                MySqlDataReader reader = ExecutQuery_Select("SELECT * FROM inwork WHERE Ser_Num = '" + serial_num + "'");
                if (reader == null) return;
                if (reader.HasRows != false) {

                    while (reader.Read()) {
                        metersTypeTB.Text = reader["TypeMeter"].ToString();
                        descriptionTB.Text = reader["Descr"].ToString();
                        producedDateTB.Text = reader.GetDateTime(7).ToString("yyyy-MM-dd hh:mm:ss");
                        analysedDateTB.Text = reader.GetDateTime(12).ToString("yyyy-MM-dd hh:mm:ss");

                        choseSolution.Text = reader["Solution"].ToString();
                        choseCustomer.Text = reader["CustomerID"].ToString();
                        choseDeliveryMethod.Text = reader["sposob_dost"].ToString();
                        mailNumberTB.Text = (DBNull.Value.Equals(reader["num_dost"])) ? "" : reader["num_dost"].ToString();
                        sendDateTB.Text = (DBNull.Value.Equals(reader["num_dost"])) ? "" : reader.GetDateTime(17).ToString("yyyy-MM-dd hh:mm:ss");
                        developmentTB.Text = reader["narabotka"].ToString();

                        withKitButton.Checked = (reader["kit"].ToString() == "1") ? true : false;
                        withoutKitButton.Checked = (reader["kit"].ToString() == "0") ? true : false;

                        defectCodeTB.Items.AddRange(reader["Codeb"].ToString().Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries));



                        addMeterButton.Text = "Обновить";

                    } 
                } else {
                    addMeterButton.Text = "Добавить новый";
                    comm.Connection.Close();

                    reader = ExecutQuery_Select("SELECT * FROM temptable WHERE Ser_Num = '" + serial_num + "'");
                    if (reader != null) {

                        while (reader.Read()) {
                            metersTypeTB.Text = reader["TypeMeter"].ToString();
                            producedDateTB.Text = reader.GetDateTime(3).ToString("yyyy-MM-dd hh:mm:ss");
                            defectCodeTB.Text = reader["Codeb"].ToString();
                            analysedDateTB.Text = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");
                        }
                    }
                }
                comm.Connection.Close();
            }
        }

        private void btn_add_pribor_Click(object sender, EventArgs e)
        {
            string codes_braka = "";
            foreach (string d in defectCodeTB.Items)
            {
                codes_braka += d + '|';
            }

            if (addMeterButton.Text.Trim() == "Обновить")
            {
                find_row = false;
                int kit = (withKitButton.Checked == true) ? 1 : 0;

                int res = ExecutQuery_Insert("UPDATE `inwork` SET `Ser_Num`='"
                    + serialNumberTB.Text + "',`TypeMeter`='"
                    + metersTypeTB.Text + "',`DateCreate`='" + producedDateTB.Text + "',`Solution`='"
                    + choseSolution.Text + "',`Codeb`='" + codes_braka + "',`Code_err_user`='" + defectCodeFromConsumer.Text + "',`CustomerID`='" + choseCustomer.Text.Split(',')[0] + "',`UserID`='"
                    + ID_User + "',`DateAnaliz`='" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "',`Descr`='" + descriptionTB.Text + "',`kit`='"
                    + kit + "',`sposob_dost`='" + choseDeliveryMethod.Text + "',`num_dost`='" + mailNumberTB.Text + "',`date_dost`='" + sendDateTB.Text
                    + "',`narabotka`='" + developmentTB.Text + "' WHERE Ser_Num='" + serialNumberTB.Text + "';");

                choseCustomer.Text = "";
                producedDateTB.Clear();
                serialNumberTB.Clear();
                metersTypeTB.Clear();
                defectCodeTB.Items.Clear();
                Refresh_DataGridView();
                developmentTB.Text = "";

            } else {

                if (chosePartysNumber.Text == "0 - Новая партия") {

                    MySqlDataReader reader = ExecutQuery_Select("SELECT MAX(IDParty) as IDParty FROM inwork");

                    if (reader != null)
                        while (reader.Read())
                        {
                            chosePartysNumber.Items[chosePartysNumber.Items.Count - 1] = Convert.ToInt32(reader["IDParty"]) + 1;
                            defectCode4FindAct.Items.Add(Convert.ToInt32(reader["IDParty"]) + 1);
                            this.chosePartysNumber.SelectedIndex = this.chosePartysNumber.Items.Count - 1;
                            //мб нужен рефреш потому что ласт итем пропадает, лечится перезагрузкой приложения
                        }
                    comm.Connection.Close();
                }

                find_row = false;
                int kit = (withKitButton.Checked == true) ? 1 : 0;

                if (choseDeliveryMethod.Text == "Непосредственно от потреб.") {
                    int res = ExecutQuery_Insert("INSERT INTO `inwork` (`IDParty`, `Ser_Num`, `TypeMeter`, `DateCreate`, `Solution`, `Codeb`, `CustomerID`, `UserID`, `DatePriem`, `DateAnaliz`, `Descr`, `kit`, `sposob_dost`, `num_dost`, `date_dost`, `narabotka`,`Code_err_user`) VALUES ('"
                        + chosePartysNumber.Text.Split('-')[0] + "', '" + serialNumberTB.Text + "', '" + metersTypeTB.Text + "', '"
                        + producedDateTB.Text + "', '" + choseSolution.Text + "', '" + codes_braka + "', '" + choseCustomer.Text.Split(',')[0]
                        + "', '" + ID_User + "', '" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "', '"
                        + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "', '" + descriptionTB.Text + "', '"
                        + kit + "', '"
                        + choseDeliveryMethod.Text + "', "
                        + "NULL, "
                        + "NULL, '"
                        + developmentTB.Text + "', '"
                        + defectCodeFromConsumer.Text + "');");
                } else {
                    int res = ExecutQuery_Insert("INSERT INTO `inwork` (`IDParty`, `Ser_Num`, `TypeMeter`, `DateCreate`, `Solution`, `Codeb`, `CustomerID`, `UserID`, `DatePriem`, `DateAnaliz`, `Descr`, `kit`, `sposob_dost`, `num_dost`, `date_dost`, `narabotka`,`Code_err_user`) VALUES ('"
                        + chosePartysNumber.Text.Split('-')[0] + "', '" + serialNumberTB.Text + "', '" + metersTypeTB.Text + "', '"
                        + producedDateTB.Text + "', '" + choseSolution.Text + "', '" + codes_braka + "', '" + choseCustomer.Text.Split(',')[0]
                        + "', '" + ID_User + "', '" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "', '"
                        + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "', '" + descriptionTB.Text + "', '"
                        + kit + "', '"
                        + choseDeliveryMethod.Text + "', '"
                        + mailNumberTB.Text + "', '"
                        + sendDateTB.Text + "', '"
                        + developmentTB.Text + "', '"
                        + defectCodeFromConsumer.Text + "');");                        
                }
                producedDateTB.Clear();
                serialNumberTB.Clear();
                metersTypeTB.Clear();
                defectCodeTB.Items.Clear();
                analysedDateTB.Text = "";
                developmentTB.Text = "";
                Refresh_DataGridView();
            }
        }

        private void btn_refresh_pribor_Click(object sender, EventArgs e)
        {
            find_row = false;

            int code = 0;
            try
            {
                code = Convert.ToInt32(defectCodeTB.Text);
            }
            catch { }
            int kit = (withKitButton.Checked == true) ? 1 : 0;
            int res = ExecutQuery_Insert("UPDATE `inwork` SET `Ser_Num`='" + serialNumberTB.Text + "',`TypeMeter`='" + metersTypeTB.Text + "',`DateCreate`='" + producedDateTB.Text + "',`Solution`='" + choseSolution.Text + "',`Codeb`='" + code + "',`CustomerID`='" + choseCustomer.Text + "',`UserID`='" + ID_User + "',`DateAnaliz`='" + analysedDateTB.Text + "',`Descr`='" + descriptionTB.Text + "',`kit`='" + kit + "' WHERE Ser_Num='" + serialNumberTB.Text + "';");

            choseCustomer.Text = "";
            producedDateTB.Clear();
            serialNumberTB.Clear();
            metersTypeTB.Clear();
            defectCodeTB.Text = "";
            choseSolution.Text = "";
            Refresh_DataGridView();

        }

        void Create_DataGridView()
        {
            var con_num = new DataGridViewColumn();
            con_num.HeaderText = "№";
            con_num.ReadOnly = true;
            con_num.Name = "Number";
            con_num.CellTemplate = new DataGridViewTextBoxCell();

            var column1 = new DataGridViewColumn();
            column1.HeaderText = "Серийный номер";
            column1.Name = "SerNum";
            column1.CellTemplate = new DataGridViewTextBoxCell();

            var column2 = new DataGridViewColumn();
            column2.HeaderText = "Тип";
            column2.Name = "Type";
            column2.CellTemplate = new DataGridViewTextBoxCell();

            var column3 = new DataGridViewColumn();
            column3.HeaderText = "Дата производства";
            column3.Name = "Date_made";
            column3.CellTemplate = new DataGridViewTextBoxCell();

            var column4 = new DataGridViewColumn();
            column4.HeaderText = "Потребитель";
            column4.Name = "Customer";
            column4.CellTemplate = new DataGridViewTextBoxCell();

            var column5 = new DataGridViewColumn();
            column5.HeaderText = "Причина";
            column5.Name = "Solution";
            column5.CellTemplate = new DataGridViewTextBoxCell();

            var column5_5 = new DataGridViewColumn();
            column5_5.HeaderText = "Код (Потреб.)";
            column5_5.Name = "CodeA";
            column5_5.CellTemplate = new DataGridViewTextBoxCell();

            var column6 = new DataGridViewColumn();
            column6.HeaderText = "Код дефекта";
            column6.Name = "CodeB";
            column6.CellTemplate = new DataGridViewTextBoxCell();

            var column7 = new DataGridViewColumn();
            column7.HeaderText = "Анализ проведен";
            column7.Name = "User";
            column7.CellTemplate = new DataGridViewTextBoxCell();

            var column8 = new DataGridViewColumn();
            column8.HeaderText = "Наработка";
            column8.Name = "Narabotka";
            column8.CellTemplate = new DataGridViewTextBoxCell();

            var column9 = new DataGridViewColumn();
            column9.HeaderText = "Дата анализа";
            column9.Name = "Date_Analiz";
            column9.CellTemplate = new DataGridViewTextBoxCell();


            var column10 = new DataGridViewColumn();
            column10.HeaderText = "Примечание";
            column10.Name = "Descr";
            column10.CellTemplate = new DataGridViewTextBoxCell();

            var column11 = new DataGridViewColumn();
            column11.HeaderText = "Комплект";
            column11.Name = "Kit";
            column11.CellTemplate = new DataGridViewTextBoxCell();

            mainDataGridView.Columns.Add(con_num);
            mainDataGridView.Columns.Add(column8);
            mainDataGridView.Columns.Add(column2);
            mainDataGridView.Columns.Add(column1);
            mainDataGridView.Columns.Add(column3);
            mainDataGridView.Columns.Add(column4);
            mainDataGridView.Columns.Add(column5);
            mainDataGridView.Columns.Add(column5_5);
            mainDataGridView.Columns.Add(column6);
            mainDataGridView.Columns.Add(column7);
            mainDataGridView.Columns.Add(column9);
            mainDataGridView.Columns.Add(column10);
            mainDataGridView.Columns.Add(column11);


            mainDataGridView.AllowUserToAddRows = false; 
            mainDataGridView.AllowUserToDeleteRows = false;
            mainDataGridView.AllowUserToOrderColumns = false;
        }
        async void Refresh_DataGridView()
        {
            if (chosePartysNumber.Text == "0 - Новая партия") return;

            mainDataGridView.Rows.Clear();

            await Task.Delay(10);
            MySqlDataReader reader = ExecutQuery_Select("SELECT * FROM `inwork` WHERE IDParty = '" + chosePartysNumber.Text + "'");
            int i = 1;

            if (reader == null) return;
            try
            {
                if (reader.HasRows != false)
                {
                    while (reader.Read())
                    {
                        mainDataGridView.Rows.Add(i, reader["narabotka"].ToString(), reader["TypeMeter"].ToString(), reader["Ser_Num"].ToString(), reader["DateCreate"].ToString(),
                            reader["CustomerID"].ToString(), reader["Solution"].ToString(), reader["Code_err_user"].ToString(), reader["Codeb"].ToString(), reader["UserID"].ToString(), reader["DateAnaliz"].ToString(), reader["Descr"].ToString(), reader["kit"].ToString());
                        i++;
                    }
                }

            }
            catch (Exception error) { MessageBox.Show(error.Message); }

            comm.Connection.Close();

            find_row = true;
            choseCustomer.Items.Clear();
            await Task.Delay(10);

            reader = ExecutQuery_Select("SELECT * FROM `customers`");

            if (reader.HasRows != false)
            {
                while (reader.Read())
                {
                    choseCustomer.Items.Add(reader["ID"].ToString() + ", " + reader["Descr"].ToString());
                }
            }

            comm.Connection.Close();

            await Task.Delay(10);
        }

        private void dataGridView1_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (find_row == false) return;

            serialNumberTB.Text = mainDataGridView["SerNum", e.RowIndex].Value.ToString();
        }

        private void Create_Priz_Nak_Click(object sender, EventArgs e)
        {
            if (defectCode4FindAct.Text.Trim() == "") return;

            MySqlDataReader reader1 = ExecutQuery_Select("SELECT * FROM inwork WHERE IDParty = '" + defectCode4FindAct.Text + "'");
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

            bookmark_NUM.Range.Text = defectCode4FindAct.Text;
            bookmark_Date.Range.Text = DateTime.Now.ToString("D");

            bookmark_NUM2.Range.Text = defectCode4FindAct.Text;
            bookmark_Date2.Range.Text = DateTime.Now.ToString("D");

            string USER_ID = "0";
            MySqlDataReader reader = ExecutQuery_Select("SELECT CustomerID, UserID FROM inwork WHERE IDParty = '" + defectCode4FindAct.Text + "'");
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

            MySqlDataReader type_data = ExecutQuery_Select("SELECT DISTINCT TypeMeter FROM inwork WHERE IDParty = '" + defectCode4FindAct.Text + "'");
            {
                if (type_data != null)
                {
                    while (type_data.Read())
                    {
                        types_meter.Add(type_data["TypeMeter"].ToString());
                    }
                }
            }
            comm.Connection.Close();

            foreach (string type in types_meter)
            {
                reasons_meter.Clear();
                MySqlDataReader reason = ExecutQuery_Select("SELECT DISTINCT Solution FROM inwork WHERE IDParty = '" + defectCode4FindAct.Text + "' AND TypeMeter = '" + type + "'");
                {
                    if (reason != null)
                    {
                        while (reason.Read())
                        {
                            reasons_meter.Add(reason["Solution"].ToString());

                        }
                    }
                }
                comm.Connection.Close();

                foreach (string solution in reasons_meter)
                {
                    MySqlDataReader withoutKit = ExecutQuery_Select("SELECT COUNT(*) as count FROM inwork WHERE IDParty = '" + defectCode4FindAct.Text + "' AND TypeMeter = '" + type + "' AND Solution = '" + solution + "'");
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
                        }
                    }
                    comm.Connection.Close();
                    MySqlDataReader withKit = ExecutQuery_Select("SELECT COUNT(*) as count FROM inwork WHERE IDParty = '" + defectCode4FindAct.Text + "' AND TypeMeter = '" + type + "' AND Solution = '" + solution + "' AND kit = '1'");
                    if (withKit != null)
                    {
                        while (withKit.Read())
                        {
                            tableWord.Cell(row - 1, 3).Range.Text = withKit["count"].ToString();
                            tableWord2.Cell(row - 1, 3).Range.Text = withKit["count"].ToString();

                            all_kit_int += Convert.ToInt32(withKit["count"]);
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


            MySqlDataReader acts = ExecutQuery_Select("SELECT COUNT(*) as count FROM acts WHERE IDParty = '" + defectCode4FindAct.Text + "'");

            if (acts != null) {
                while (acts.Read())
                {
                    if (Convert.ToInt32(acts["count"]) > 0) {
                        comm.Connection.Close();
                    } else {
                        comm.Connection.Close();
                        Form_acts(defectCode4FindAct.Text);

                    }
                    break;
                }
            } else { MessageBox.Show("Записей нет"); }

            comm.Connection.Close();

            if (withActsCheckBox.Checked) { Show_All_Acts(Convert.ToInt32(defectCode4FindAct.Text)); }
        }

        void fill_cb()
        {
            MySqlDataReader reader = ExecutQuery_Select("SELECT * FROM reasonreturn");
            if (reader != null)
            {
                while (reader.Read())
                {
                    choseSolution.Items.Add(reader["Reason"].ToString());
                }
            }
            comm.Connection.Close();
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e) { }

        void FillDataGridViewCustomer()
        {
            agentsDataGridView.Rows.Clear();

            MySqlDataReader reader = ExecutQuery_Select("SELECT * FROM `customers`");

            int i = 1;

            if (reader == null) return;
            if (reader.HasRows != false) {
                while (reader.Read())
                {
                    agentsDataGridView.Rows.Add(reader["ID"].ToString(), reader["Descr"].ToString(), reader["name"], reader["surname"], reader["address"], reader["ContFace"].ToString(), reader["Phone"].ToString(), reader["_Index"].ToString(),
                        reader["Resp"].ToString(), reader["Oblast"].ToString(), reader["City"].ToString(), reader["Street"].ToString(), reader["Num_h"].ToString(), reader["Num_f"].ToString());
                    i++;
                }
            }

            comm.Connection.Close();
        }

        private void cust_id_TextChanged(object sender, EventArgs e)
        {
            add_customer.Text = (cust_id.Text.Trim() == "") ? "Добавить нового" : "Обновить";

            if (cust_id.Text.Trim() != "") {
                MySqlDataReader reader = ExecutQuery_Select("SELECT * FROM `customers` WHERE ID = '" + cust_id.Text.Trim() + "'");

                if (reader == null) return;
                if (reader.HasRows != false) {
                    while (reader.Read())
                    {
                        cust_name.Text = reader["Descr"].ToString();
                        customerNameTextBox.Text = reader["name"].ToString();
                        customerLastnameTextBox.Text = reader["surname"].ToString();
                        customerAddressTextBox.Text = reader["address"].ToString();
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
             
        async private void add_customer_Click(object sender, EventArgs e)
        {
            if (cust_id.Text.Trim() != "")
            {
                int res = ExecutQuery_Insert("UPDATE `customers` SET `Descr`='" + cust_name.Text + "" +
                    "',`ContFace`='" + cust_face.Text + "" +
                    "',`name`='" + customerNameTextBox.Text + "" +
                    "',`surname`='" + customerLastnameTextBox.Text + "" +
                    "',`address`='" + customerAddressTextBox.Text + "" +
                    "',`Phone`='" + cust_phone.Text + "',`_Index`='" + cust_index.Text +
                    "',`Resp`='" + cust_resp.Text + "',`Oblast`='" + cust_raion.Text +
                    "',`City`='" + cust_city.Text + "',`Street`='" + cust_street.Text +
                    "',`Num_h`='" + cust_house.Text + "',`Num_f`='" + cust_flat.Text + "' WHERE ID='" + cust_id.Text + "';");
            }
            else
            {
                int res = ExecutQuery_Insert("INSERT INTO `customers` (`ID`, `name`, `surname`, `address`, `Descr`, `ContFace`, `Phone`, `_Index`, `Resp`, `Oblast`, `City`, `Street`, `Num_h`, `Num_f`) " +
                    "VALUES (NULL, '" + customerNameTextBox.Text + "', '" + customerLastnameTextBox.Text + "', '" + customerAddressTextBox.Text + "', '" + cust_name.Text + "', '" + cust_face.Text + "', '" + cust_phone.Text + "', '" + cust_index.Text + "', '"
                    + cust_resp.Text + "', '" + cust_raion.Text + "', '" + cust_city.Text + "', '" + cust_street.Text + "', '" + cust_house.Text + "', '" + cust_flat.Text + "');");
            }

            await Task.Delay(10);
            FillDataGridViewCustomer();
            Refresh_DataGridView();
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
            if (choseDeliveryMethod.Text == "Непосредственно от потреб.")
            {
                label24.Visible = false;
                label25.Visible = false;
                mailNumberTB.Visible = false;
                sendDateTB.Visible = false;
                sendDateTB.Text = "";
                mailNumberTB.Text = "";
            }
            else
            {
                label24.Visible = true;
                label25.Visible = true;
                mailNumberTB.Visible = true;
                sendDateTB.Visible = true;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string query = "SELECT * FROM inwork WHERE (";
            string text_data = "Отчет по \"";
            bool have_meter = false;
            bool have_code_err = false;
            string[] errors = new string[defectCodesList.Items.Count];
            int error_index = 0;

            for (int i = 0; i < meterTypesList.Items.Count; i++)
            {
                if (meterTypesList.GetItemChecked(i)) {
                    if (have_meter == true) {
                        query += " OR ";
                    }

                    have_meter = true;
                    query += "TypeMeter = '" + meterTypesList.Items[i] + "'";
                    text_data += meterTypesList.Items[i] + " ";
                }
            }

            if (have_meter == false) { query += "1"; }

            query += ") AND (";
            text_data += "\" по \"";
            have_meter = false;
            for (int i = 0; i < defectCodesList.Items.Count; i++)
            {
                if (defectCodesList.GetItemChecked(i)) {
                    if (have_meter == true) { query += " OR "; }

                    have_meter = true;
                    query += "Codeb LIKE '%" + defectCodesList.Items[i] + "%'";
                    text_data += defectCodesList.Items[i] + " ";
                    have_code_err = true;

                    errors[error_index] = defectCodesList.Items[i] + "";
                    error_index++;
                }
            }

            if (have_meter == false) { query += "1"; }

            query += ") AND (DatePriem BETWEEN ";
            text_data += "\" за ";

            if (choseType4Excel.Checked == true) {
                text_data += reportYearStart.Text + "|" + reportQuarterStart.Text + "| : " + reportYearEnd.Text + "|" + reportQuarterEnd.Text + "|";
                if (reportQuarterStart.Text == reportQuarterEnd.Text && reportYearEnd.Text == reportYearStart.Text)
                {
                    switch (reportQuarterStart.Text)
                    {
                        case "1": query += "'" + reportYearStart.Text + "-" + "01" + "-" + "01 00:00:00' and '" + reportYearStart.Text + "-" + "03" + "-" + "31 23:59:59'"; break;
                        case "2": query += "'" + reportYearStart.Text + "-" + "04" + "-" + "01 00:00:00' and '" + reportYearStart.Text + "-" + "06" + "-" + "30 23:59:59'"; break;
                        case "3": query += "'" + reportYearStart.Text + "-" + "07" + "-" + "01 00:00:00' and '" + reportYearStart.Text + "-" + "09" + "-" + "30 23:59:59'"; break;
                        case "4": query += "'" + reportYearStart.Text + "-" + "10" + "-" + "01 00:00:00' and '" + reportYearStart.Text + "-" + "12" + "-" + "31 23:59:59'"; break;
                        default: query += "'" + reportYearStart.Text + "-" + "01" + "-" + "01 00:00:00' and '" + reportYearStart.Text + "-" + "03" + "-" + "31 23:59:59'"; break;
                    }
                }
                else
                {
                    switch (reportQuarterStart.Text)
                    {
                        case "1": query += "'" + reportYearStart.Text + "-" + "01" + "-" + "01 00:00:00'"; break;
                        case "2": query += "'" + reportYearStart.Text + "-" + "04" + "-" + "01 00:00:00'"; break;
                        case "3": query += "'" + reportYearStart.Text + "-" + "07" + "-" + "01 00:00:00'"; break;
                        case "4": query += "'" + reportYearStart.Text + "-" + "10" + "-" + "01 00:00:00'"; break;
                        default: query += "'" + reportYearStart.Text + "-" + "01" + "-" + "01 00:00:00'"; break;
                    }
                    query += " and ";
                    switch (reportQuarterEnd.Text)
                    {
                        case "1": query += "'" + reportYearEnd.Text + "-" + "03" + "-" + "31 23:59:59'"; break;
                        case "2": query += "'" + reportYearEnd.Text + "-" + "06" + "-" + "30 23:59:59'"; break;
                        case "3": query += "'" + reportYearEnd.Text + "-" + "09" + "-" + "30 23:59:59'"; break;
                        case "4": query += "'" + reportYearEnd.Text + "-" + "12" + "-" + "31 23:59:59'"; break;
                        default: query += "'" + reportYearEnd.Text + "-" + "03" + "-" + "31 23:59:59'"; break;
                    }
                }
            }
            else
            {
                if (dateFromTB.Text == "" || dateToTB.Text == "") { MessageBox.Show("Пожалуйста, укажите временной интервал ОТ и ДО при помощи календаря!"); return; }

                text_data += dateFromTB.Text + " : " + dateToTB.Text;
                query += "'" + dateFromTB.Text + " 00:00:00' and '" + dateToTB.Text + " 23:59:59'";
            }

            query += ")";


            richTextBox1.Text = query;

            string[] rimsk = new string[] { "0", "I", "II", "III", "IV" };

            MySqlDataReader data = ExecutQuery_Select(query);

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

                int year = Convert.ToInt32(reportYearStart.Text);
                int kv = Convert.ToInt32(reportQuarterStart.Text);
                if (have_code_err == false)
                {
                    if (choseType4Excel.Checked == true)
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

                        while (year <= Convert.ToInt32(reportYearEnd.Text))
                        {
                            sheet.Cells[row, 1] = rimsk[kv] + " кв. " + year + " г.";
                            row++;
                            kv++;
                            if (kv - 1 == Convert.ToInt32(reportQuarterEnd.Text) && year == Convert.ToInt32(reportYearEnd.Text))
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

                    if (choseType4Excel.Checked == true)
                    {
                        int years = Convert.ToInt32(reportYearEnd.Text) - Convert.ToInt32(reportYearStart.Text);
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

                        int start_year = Convert.ToInt32(reportYearStart.Text);
                        int end_year = Convert.ToInt32(reportYearEnd.Text);

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

                                var dataCodeB = (sheet.Cells[new_row, 2] as Excel.Range)?.Value2?.ToString() ?? "";

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
            cust_id.Text = agentsDataGridView["Column1", e.RowIndex].Value.ToString();
        }

        void Form_acts(string act_IDParty)
        {
            MySqlDataReader max_id = ExecutQuery_Select("SELECT MAX(ID) as Max_ID FROM acts");

            int act_ID = 0;
            string act_Date = "", act_sposob_dost = "", act_num_dost = null, act_date_dost = null, customerid = "";

            if (max_id != null) {
                while (max_id.Read()) {
                    if (max_id["Max_ID"].ToString().Trim() != "") {
                        act_ID = Convert.ToInt32(max_id["Max_ID"].ToString());
                    }
                }
            }

            act_ID++;
            comm.Connection.Close();

            MySqlDataReader find_date = ExecutQuery_Select("SELECT * FROM inwork WHERE IDParty='" + act_IDParty + "' ORDER BY DatePriem DESC");
            MySqlDataReader selectedIdCustomer = ExecutQuery_Select("SELECT * FROM inwork WHERE IDParty='" + act_IDParty + "' ORDER BY DatePriem DESC");

            if (find_date != null)
            {
                while (find_date.Read())
                {

                    act_Date = find_date["DatePriem"].ToString();
                    act_sposob_dost = find_date["sposob_dost"].ToString();
                    act_num_dost = Convert.IsDBNull(find_date["num_dost"]) ? "NULL" : find_date["num_dost"].ToString();
                    act_date_dost = Convert.IsDBNull(find_date["date_dost"]) ? "NULL" : find_date["date_dost"].ToString();
                    customerid = find_date["CustomerID"].ToString();
                    break;
                }
            }

            comm.Connection.Close();

            List<string> reasons = new List<string>();
            List<string> like_types_meters = new List<string>() {
                "(TypeMeter LIKE '%СХВ%' OR TypeMeter LIKE '%СГВ%')",
                "(TypeMeter LIKE '%СГБМ%' OR TypeMeter LIKE '%СГБУ%')",
                "(TypeMeter LIKE '%СВМ%')",
                "(TypeMeter LIKE '%ЭСО%' OR TypeMeter LIKE '%БАРС%')",
                "(TypeMeter LIKE '%СТК%')",
                "(TypeMeter LIKE '%ПВМ%')",
                "(TypeMeter LIKE '%РД%')" 
            };

            MySqlDataReader find_reasons = ExecutQuery_Select("SELECT * FROM reasonreturn");

            if (find_reasons != null) {
                while (find_reasons.Read()) {
                    reasons.Add(find_reasons["Reason"].ToString().Trim());
                }
            }

            comm.Connection.Close();
            List<meter> meters = new List<meter>();

            foreach (string type_meter in like_types_meters) {
                foreach (string reason in reasons) {

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

                    if (find_meter != null) {
                        while (find_meter.Read()) {
                            meters.Add(new meter() {
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

                        if (act_date_dost == "NULL"  || act_num_dost == "NULL") {
                            ExecutQuery_Insert("INSERT INTO `acts`" +
                                "(`ID`, `Date`, `IDParty`," +
                                " `sposob_dost`, `num_dost`, `date_dost`," +
                                " `ID_Customer`, `Meters`, `Solution`, `TypeMeters`, `ID_User`)" +
                                " VALUES ('" + act_ID + "', '" + Convert.ToDateTime(meters[0].DatePriem).ToString("yyyy-MM-dd hh:mm:ss") + "', '" + act_IDParty +
                                "', '" + act_sposob_dost + "', NULL, NULL" +
                                ", '" + customerid + "', '" + data_meter_json + "', '" + reason + "', '" + type_prib + "','" + meters[0].UserID + "')");
                        } else {
                            ExecutQuery_Insert("INSERT INTO `acts`" +
                                "(`ID`, `Date`, `IDParty`," +
                                " `sposob_dost`, `num_dost`, `date_dost`," +
                                " `ID_Customer`, `Meters`, `Solution`, `TypeMeters`, `ID_User`)" +
                                " VALUES ('" + act_ID + "', '" + Convert.ToDateTime(meters[0].DatePriem).ToString("yyyy-MM-dd hh:mm:ss") + "', '" + act_IDParty +
                                "', '" + act_sposob_dost + "', '" + act_num_dost + "', '" + Convert.ToDateTime(act_date_dost).ToString("yyyy-MM-dd hh:mm:ss") + "'," +
                                "'" + customerid + "', '" + data_meter_json + "', '" + reason + "', '" + type_prib + "','" + meters[0].UserID + "')");
                        }
                        act_ID++;
                    }

                    comm.Connection.Close();
                }
            }
        }
        void Show_All_Acts(int IDParty)
        {
            List<int> id_acts_int = new List<int>();
            MySqlDataReader id_acts = ExecutQuery_Select("SELECT * FROM acts WHERE IDParty = '" + IDParty + "'");

            if (id_acts != null) {
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

            if (reader1 != null) {
                if (!reader1.HasRows) {
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

            string bd_ID_act = "", bd_date_act = "", bd_IDParty_act = "", bd_sp_dost_act = "", bd_num_dost_act = "", bd_date_dost_act = "";
            string bd_ID_Customer = "", bd_Solution_act = "", bd_ID_User_act = "", bd_ID_Meters = "";
            string[] metersID = new string[] { };

            MySqlDataReader readact = ExecutQuery_Select("SELECT * FROM acts WHERE ID = '" + ID_Act + "'");

            if (readact == null) return;

            try
            {
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
                        bd_ID_Customer = readact["ID_Customer"].ToString();
                        bd_ID_Meters = readact["Meters"].ToString();
                        bd_Solution_act = readact["Solution"].ToString();
                        bd_ID_User_act = readact["ID_User"].ToString();
                    }
                }
            }
            catch (Exception error) { MessageBox.Show(error.Message);}

            comm.Connection.Close();

            bookmark_date_act.Range.Text = bookmark_date_act2.Range.Text = bookmark_date_act3.Range.Text = bd_date_act.Substring(0, 10);
            bookmark_num_act.Range.Text = bd_ID_act;
            bookmark_num_pn_act.Range.Text = bd_IDParty_act;
            bookmark_sp_pol_act.Range.Text = bd_sp_dost_act;
            bookmark_sp_pol_num_act.Range.Text = bd_num_dost_act;
            bookmark_sp_pol_date_act.Range.Text = bd_date_dost_act;

            MySqlDataReader customer = ExecutQuery_Select("SELECT * FROM customers WHERE " + bd_ID_Customer + " = ID");
            try
            {
                while (customer.Read())
                {
                    bookmark_customer_act.Range.Text = customer["Descr"].ToString();
                    bookmark_cust_index_act.Range.Text = customer["_Index"].ToString();
                    bookmark_cust_resp_act.Range.Text = customer["Resp"].ToString();
                    bookmark_cust_oblast_act.Range.Text = customer["Oblast"].ToString();
                    bookmark_cust_city_act.Range.Text = customer["City"].ToString();
                    bookmark_cust_street_act.Range.Text = customer["Street"].ToString();
                    bookmark_cust_h_act.Range.Text = customer["Num_h"].ToString();
                    bookmark_cust_flat_act.Range.Text = customer["Num_f"].ToString();
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); return; }

            comm.Connection.Close();

            if (bd_Solution_act == "Негарантия") bookmark_negarant_act.Range.Text = "V";

            MySqlDataReader readusr = ExecutQuery_Select("SELECT * FROM users WHERE ID = '" + bd_ID_User_act + "'");

            if (readusr.HasRows != false) {
                while (readusr.Read())
                {
                    bookmark_user_act.Range.Text = readusr["Descr"].ToString();
                }
            }
            comm.Connection.Close();

            bookmark_solution_act.Range.Text = bd_Solution_act;
            int all_prib_without_kit = 0, all_prib_with_kit = 0, row = 2,column = 3;

            meter[] meters = JsonConvert.DeserializeObject<meter[]>(bd_ID_Meters);

            for (int index = 0; index < meters.Length; index++)
            {
                int sum_all = 0;
                int sum_kompl = 0;
                string type_meter_cycle = meters[index].TypeMeter;

                if (index == meters.Length - 1)
                {
                    if (index < 10)
                    {
                        row = 10;
                        column = 6;
                    }
                    docWord.Tables[2].Cell(row, column).Range.Text = "Всего:";
                    docWord.Tables[2].Cell(row, column + 1).Range.Text = all_prib_without_kit.ToString();
                    docWord.Tables[2].Cell(row, column + 2).Range.Text = all_prib_with_kit.ToString();
                }

                if (type_meter_cycle == "none") continue;

                for (int a = 0; a < meters.Length; a++)
                {
                    if (meters[a].TypeMeter == type_meter_cycle)
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

                docWord.Tables[2].Cell(row, column).Range.Text = type_meter_cycle;
                docWord.Tables[2].Cell(row, column + 1).Range.Text = sum_all.ToString();
                docWord.Tables[2].Cell(row, column + 2).Range.Text = sum_kompl.ToString();
                row++;

                if (row >= 10) docWord.Tables[2].Rows.Add(ref missobj);
            }

            comm.Connection.Close();

            Word.Bookmark[] bms = new Word.Bookmark[] { bookmark_ser_prib1_act, bookmark_ser_prib2_act, bookmark_ser_prib3_act, bookmark_ser_prib4_act, bookmark_ser_prib5_act };

            if (all_prib_without_kit <= 5)
            {
                for (int i = 0; i < meters.Length; i++)
                {
                    bms[i].Range.Text = meters[i].Ser_Num;
                }
            }
        }

        private void act_priem_num_acta_Click(object sender, EventArgs e)
        {
            receiversActNumberTB.Items.Clear();
            if (receiverPNactsTB.Text.Trim() == "") return;

            MySqlDataReader find_acts = ExecutQuery_Select("SELECT * FROM acts WHERE IDParty = '" + receiverPNactsTB.Text + "'"); ;
            if (find_acts == null) return;
            if (find_acts.HasRows != false)
            {
                while (find_acts.Read())
                {
                    receiversActNumberTB.Items.Add(find_acts["ID"]);
                }
            }
            comm.Connection.Close();
        }

        private void Find_Act_Click(object sender, EventArgs e)
        {
            try
            {
                Show_Act(Convert.ToInt32(receiversActNumberTB.Text));
            }
            catch (Exception ex) { MessageBox.Show(ex.Message);}
        }

        private void cb_coded_Click(object sender, EventArgs e)
        {
            if (defectCodeTB.Text.Trim() == "") {
                defectCodeTB.Items.Clear();
                defectCodeTB.Items.AddRange(code_err_water);
            }
        }

        private void tb_CLA_num_act_TextChanged(object sender, EventArgs e)
        {
            string num_act = controlListViewTB.Text;
            label30.Text = label32.Text = "";
            MySqlDataReader act = ExecutQuery_Select("SELECT * FROM acts WHERE ID = '" + num_act + "'");

            if (act != null) {
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
            string num_act = controlListViewTB.Text;
            MySqlDataReader reader1 = ExecutQuery_Select("SELECT * FROM acts WHERE ID = '" + num_act + "'");

            if (reader1 != null) {
                if (!reader1.HasRows) {
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

            string bd_schitciki_JSON = "", bd_customer_JSON = "", customerDescription = "";

            MySqlDataReader actRead = ExecutQuery_Select("SELECT * FROM acts WHERE ID = '" + num_act + "'");

            if (actRead != null) {
                while (actRead.Read()) {
                    bd_schitciki_JSON = actRead["Meters"].ToString();
                    bd_customer_JSON = actRead["ID_Customer"].ToString();
                }
            }

            comm.Connection.Close();

            MySqlDataReader customerRead = ExecutQuery_Select("SELECT * FROM customers WHERE ID = '" + bd_customer_JSON + "'");

            try
            {
                if (customerRead != null)
                {
                    while (customerRead.Read())
                    {
                        customerDescription = customerRead["Descr"].ToString();
                    }
                }
            }
            catch (Exception error) { MessageBox.Show(error.Message);}

            bookmark_ref_customer.Range.Text = customerDescription;
            bookmark_ref_num.Range.Text = num_act;

            meter[] meters = JsonConvert.DeserializeObject<meter[]>(bd_schitciki_JSON);

            foreach (meter my_meter in meters)
            {
                if (row == 2) bookmark_ref_date_an.Range.Text = my_meter.DateAnaliz.Substring(0, 10);
                if (row == 16) {
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

        private void btn_showCD_Click(object sender, EventArgs e)
        {
            Chose_CD _CD = new Chose_CD(code_err_water, ref defectCodeTB);
            _CD.ShowDialog();
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            meterTypesList.Items.Clear();

            if (tabControl1.SelectedIndex == 1) {

                MySqlDataReader reader1 = ExecutQuery_Select("SELECT TypeMeter FROM `inwork` GROUP BY TypeMeter");

                if (reader1 != null)
                {
                    if (reader1.HasRows)
                    {
                        while (reader1.Read())
                        {
                            meterTypesList.Items.Add(reader1["TypeMeter"].ToString());
                        }
                    }
                }
                comm.Connection.Close();
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (choseType4Excel.Checked == true) {
                reportQuarterEnd.Enabled = reportQuarterStart.Enabled = reportYearEnd.Enabled = reportYearStart.Enabled = true;
                dateFromTB.Enabled = dateToTB.Enabled = monthCalendarTo.Enabled = monthCalendarFrom.Enabled = false;
            } else {
                reportQuarterEnd.Enabled = reportQuarterStart.Enabled = reportYearEnd.Enabled = reportYearStart.Enabled = false;
                dateFromTB.Enabled = dateToTB.Enabled = monthCalendarTo.Enabled = monthCalendarFrom.Enabled = true;
            }
        }

        private void monthCalendar1_DateChanged(object sender, DateRangeEventArgs e)
        {
            dateFromTB.Text = e.Start.ToString("yyyy-MM-dd");
            monthCalendarTo.MinDate = e.Start;
        }

        private void monthCalendar2_DateChanged(object sender, DateRangeEventArgs e) { dateToTB.Text = e.Start.ToString("yyyy-MM-dd"); }

        private void Num_Party_SelectedIndexChanged(object sender, EventArgs e) { Refresh_DataGridView(); }
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
        public string Meters { get; set; }
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
        public string name { get; set; }
        public string lastname { get; set; }
        public string address{ get; set; }
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
