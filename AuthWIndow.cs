using MySql.Data.MySqlClient;
using System;
using System.ComponentModel;
using System.Data;
using System.Security.Principal;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;

namespace Garant1._0
{
    public partial class AuthWIndow : Form
    {
        public MySqlCommand comm;

        public AuthWIndow()
        {
            InitializeComponent();

            //HostTextBox.Text = "127.0.0.1";
            //LoginTextBox.Text = "root";
            //PasswordTextBox.Text = "";
            //DBTextBox.Text = "Garant";
            //PortTextBox.Text = "3306";

            var role = GetCurrentRole();
            this.Text = this.Text + $"\t\t\t\t\t\t\t Run As {role}";
            CheckAuthSQL();
        }   

        void CheckAuthSQL()
        {
            MySqlDataReader query = ExecuteSQL("SELECT COUNT(*) FROM connection WHERE 1;"); 
            if (query != null)
                while (query.Read())
                {
                    HostTextBox.Text = query.GetString(1);
                    DBTextBox.Text = query.GetString(2);
                    LoginTextBox.Text = query.GetString(3);
                    PasswordTextBox.Text = query.GetString(4);
                }
            comm.Connection.Close();

            MessageBox.Show(query.ToString());
        }

        MySqlDataReader ExecuteSQL(String query)
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
                MessageBox.Show(ex.Message);
            }
            return reader;
        }

        private void ConnectButton_Click(object sender, EventArgs e)
        {
            try
            {
                string server = string.Empty, login = string.Empty, password = string.Empty, database = string.Empty;

                server = HostTextBox.Text;
                login = LoginTextBox.Text;
                password = PasswordTextBox.Text;
                database = DBTextBox.Text;

                SQLConnect(server, database, login, password);
            }
            catch (Exception error) { MessageBox.Show(error.Message);}
        }
        public void SQLConnect(string server, string db, string login, string password)
        {
            MySqlCommand comm = new MySqlCommand();
            //string connect = "server=" + server + ";user=" + login + ";database=" + db + ";port=3306;password=" + password + ";convert zero datetime=True";
            string connect = "server=" + server + ";user=" + login + ";database=" + db + ";port=3306;password=" + password + ";charset=utf8mb4;convert zero datetime=True";

            MySqlConnection connection = new MySqlConnection(connect);
            comm.Connection = connection;

            try
            {
                comm.Connection.Open();

                Form1 mainForm = new Form1(comm);
                this.Hide();
                mainForm.ShowDialog();
                Application.Exit();
            }
            catch (MySql.Data.MySqlClient.MySqlException error) 
            { 
                switch (error.Number)
                {
                    case 0:
                        MessageBox.Show("База данных не найдена.\nПроверьте правильность введённых данных или обратитесь к системному администратору.");
                        break;
                    case 1042:
                        MessageBox.Show("Невозможно подключиться к серверу по указанному адресу.\nПроверьте правильность введённых данных или обратитесь к системному администратору.");
                        break;
                    default:
                        MessageBox.Show(error.Number.ToString());
                        break;
                }
            }
        }
        public string GetCurrentRole()
        {
            string role = string.Empty;
            try
            {
                WindowsIdentity user = WindowsIdentity.GetCurrent();
                WindowsPrincipal principal = new WindowsPrincipal(user);
                role = principal.IsInRole(WindowsBuiltInRole.Administrator) == true ? "Administrator" : principal.IsInRole(WindowsBuiltInRole.User) == true ? "User" : "Unrecognizned";
            }
            catch (UnauthorizedAccessException ex) { MessageBox.Show(ex.Message);}
            catch (Exception ex) { }

            return role;
        }
    }
}
