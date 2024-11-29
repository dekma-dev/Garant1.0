using MySql.Data.MySqlClient;
using System;
using System.Data.SqlClient;
using System.Security.Principal;
using System.Windows.Forms;

namespace Garant1._0
{
    public partial class AuthWIndow : Form
    {
        public MySqlCommand comm;

        public AuthWIndow()
        {
            InitializeComponent();

            HostTextBox.Text = Properties.Settings.Default.Host;
            LoginTextBox.Text = Properties.Settings.Default.Login;
            PasswordTextBox.Text = Properties.Settings.Default.Password;
            DBTextBox.Text = Properties.Settings.Default.DB;

            var role = GetCurrentRole();
            this.Text = this.Text + $"\t\t\t\t\t\t\t Run As {role}";
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

            try
            {
                MySqlConnection connection = new MySqlConnection(connect);
                comm.Connection = connection;
                connection.Open();

                Properties.Settings.Default.Host = HostTextBox.Text;
                Properties.Settings.Default.Login = LoginTextBox.Text;
                Properties.Settings.Default.Password = PasswordTextBox.Text;
                Properties.Settings.Default.DB = DBTextBox.Text;
                Properties.Settings.Default.Save();
            }
            catch (MySql.Data.MySqlClient.MySqlException error) 
            { 
                switch (error.Number)
                {
                    case 0:
                        MessageBox.Show("Неверный SQL запрос к базе данных.\nПроверьте правильность введённых данных или обратитесь к системному администратору.");
                        break;
                    case 1042:
                        MessageBox.Show("Невозможно подключиться к серверу по указанному адресу.\nПроверьте правильность введённых данных или обратитесь к системному администратору.");
                        break;
                    default:
                        MessageBox.Show(error.Number.ToString());
                        break;
                }
            }
            if (comm.Connection.State == System.Data.ConnectionState.Open)
            {
                Form1 mainForm = new Form1(ref comm);
                this.Hide();
                mainForm.ShowDialog();
                Application.Exit();
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
            catch (Exception ex) { MessageBox.Show(ex.Message);}

            return role;
        }
    }
}
