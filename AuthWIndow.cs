using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Garant1._0
{
    public partial class AuthWIndow : Form
    {
        public MySqlCommand comm;

        public AuthWIndow()
        {
            InitializeComponent();

            HostTextBox.Text = "127.0.0.1";
            LoginTextBox.Text = "root";
            PasswordTextBox.Text = "";
            DBTextBox.Text = "Garant";
            PortTextBox.Text = "3306";
        }

        private void ConnectButton_Click(object sender, EventArgs e)
        {
            try
            {
                string server = null, login = null, password = null, database = null, port = null;

                server = HostTextBox.Text;
                login = LoginTextBox.Text;
                password = PasswordTextBox.Text;
                database = DBTextBox.Text;
                port = PortTextBox.Text;

                SQLConnect(server, port, database, login, password);
            }
            catch (Exception error) { MessageBox.Show(error.Message);}
        }
        public void SQLConnect(string server = "127.0.0.1", string port = "3306", string db = "Garant", string login = "root", string password = "")
        {
            MySqlCommand comm = new MySqlCommand();
            string connect = "server=" + server + ";user=" + login + ";database=" + db + ";port=" + port + ";password=" + password + ";charset=utf8mb4;convert zero datetime=True";

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
                        MessageBox.Show(error.Message);
                        break;
                    default:
                        MessageBox.Show("Проверьте введённые данные");
                        break;
                }
            }
        }
    }
}
