using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ASAP_Project
{
    public partial class LoginScreen : Form
    {
        
        private string password = "1234"; //herkes için
        private string username = "esat"; //herkes için
        private string admin_password = "admin"; //admin için
        private string admin_username = "admin"; //admin için
        public LoginScreen()
        {
            InitializeComponent();
            panel1.BackColor = Color.FromArgb(25, Color.Black);
            panel2.BackColor = Color.FromArgb(25, Color.Black);
            panel3.BackColor = Color.FromArgb(25, Color.Black);
            panel4.BackColor = Color.FromArgb(25, Color.Black);
            panel5.BackColor = Color.FromArgb(25, Color.Black);
            pictureBox1.BackColor = Color.FromArgb(25, Color.Black);
            label_username.BackColor = Color.Transparent;
            label_password.BackColor = Color.Transparent;
            pictureBox1.BackColor = Color.Transparent;
        }

        private void LoginScreen_Load(object sender, EventArgs e)
        {

        }

        private void button_login_Click(object sender, EventArgs e)
        {
            if(textBox_password.Text== password &&  textBox_username.Text== username || textBox_password.Text == admin_password && textBox_username.Text == admin_username)
            {
                MainPage form = new MainPage();
                form.Show();

                this.Hide();
            }
            else
                MessageBox.Show("Incorrect Usarname or Password");

        }
    }
}
