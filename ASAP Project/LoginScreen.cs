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

    }
}
