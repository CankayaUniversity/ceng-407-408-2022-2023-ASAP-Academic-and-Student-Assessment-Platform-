using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System;

using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.Logging;
using Word = Microsoft.Office.Interop.Word;

namespace ASAP_Project
{
    public partial class MainPage : Form
    {
        [DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]

        private static extern IntPtr CreateRoundRectRgn
        (
            int nLeftRect,
            int nTopRect,
            int nRightRect,
            int nBottomRect,
            int nWidthEllipse,
            int nHeightEllipse
        );
        public void MainPage_Load()
        {
            panel_adminpanel.Visible = false;
            panel_userpanel.Visible = false;
            panel1.BackColor = Color.Transparent;
            panel2.BackColor = Color.Transparent;
            panel3.BackColor = Color.Transparent;
            panel4.BackColor = Color.Transparent;
            panel5.BackColor = Color.FromArgb(60, Color.Black);
            panel_userpanel.BackColor = Color.FromArgb(60, Color.Black);
            panel_adminpanel.BackColor = Color.FromArgb(60, Color.Black);
            panel5.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, panel5.Width,
            panel5.Height, 30, 30));
            panel_userpanel.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, panel_userpanel.Width,
            panel_userpanel.Height, 30, 30));
            panel_adminpanel.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, panel_adminpanel.Width,
            panel_adminpanel.Height, 30, 30));
            pictureBox1.BackColor = Color.Transparent;
            label_user.Text += LoginScreen.user_name;
            label_user.BackColor = Color.Transparent;

            button1.BackColor = Color.FromArgb(90, Color.Black);
            button1.FlatAppearance.MouseOverBackColor = Color.FromArgb(110, Color.Black);
            button1.FlatAppearance.MouseDownBackColor = Color.FromArgb(130, Color.Black);
            button1.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, button1.Width,
            button1.Height, 5, 5));
            button1.ForeColor = Color.LightGray;

            button_userpanel.BackColor = Color.FromArgb(90, Color.Black);
            button_userpanel.FlatAppearance.MouseOverBackColor = Color.FromArgb(110, Color.Black);
            button_userpanel.FlatAppearance.MouseDownBackColor = Color.FromArgb(130, Color.Black);
            button_userpanel.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, button_userpanel.Width,
            button_userpanel.Height, 5, 5));
            button_userpanel.ForeColor = Color.LightGray;

            button_adminpanel.BackColor = Color.FromArgb(90, Color.Black);
            button_adminpanel.FlatAppearance.MouseOverBackColor = Color.FromArgb(110, Color.Black);
            button_adminpanel.FlatAppearance.MouseDownBackColor = Color.FromArgb(130, Color.Black);
            button_adminpanel.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, button_adminpanel.Width,
            button_adminpanel.Height, 5, 5));
            button_adminpanel.ForeColor = Color.LightGray;

            button_account.BackColor = Color.FromArgb(90, Color.Black);
            button_account.FlatAppearance.MouseOverBackColor = Color.FromArgb(110, Color.Black);
            button_account.FlatAppearance.MouseDownBackColor = Color.FromArgb(130, Color.Black);
            button_account.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, button_account.Width,
            button_account.Height, 5, 5));
            button_account.ForeColor = Color.LightGray;

            button_exit.BackColor = Color.FromArgb(90, Color.Black);
            button_exit.FlatAppearance.MouseOverBackColor = Color.FromArgb(110, Color.Black);
            button_exit.FlatAppearance.MouseDownBackColor = Color.FromArgb(130, Color.Black);
            button_exit.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, button_exit.Width,
            button_exit.Height, 5, 5));
            button_exit.ForeColor = Color.LightGray;
        }

        public MainPage()
        {
            InitializeComponent();
            MainPage_Load();

        }

        

        private void button_userpanel_Click(object sender, EventArgs e)
        {

            panel_adminpanel.Visible = false;
            panel_adminpanel.Enabled = false;
            panel_userpanel.Enabled = true;
            panel_userpanel.Visible = true;
            panel_userpanel.BringToFront();

        }

        private void button_adminpanel_Click(object sender, EventArgs e)
        {
            if (LoginScreen.user_name == "admin" && LoginScreen.user_password == "admin")
            {
                panel_adminpanel.Visible = true;
                panel_adminpanel.Enabled = true;
            }
            else
            {
                MessageBox.Show("You need login as admin!");
                panel_adminpanel.Enabled = false;
                panel_adminpanel.Visible = false;
            }
            panel_userpanel.Visible = false;
            panel_userpanel.Enabled = false;

            panel_adminpanel.BringToFront();
        }

        private void button_account_Click(object sender, EventArgs e)
        {

        }

        private void button_testdrive_Click(object sender, EventArgs e)
        {
            try
            {
                ASAP_Project.GoogleDrive.UploadFile();

            }
            catch (Exception ex)
            {
                throw (ex);
                MessageBox.Show("Error");
            }
        }

        private void button_exit_Click(object sender, EventArgs e)
        {

            DialogResult secenek = MessageBox.Show("Çýkýþ yapmak istediðinize emin misiniz?", "Dikkat", MessageBoxButtons.YesNo, MessageBoxIcon.Information);

            if (secenek == DialogResult.Yes)
            {
                Environment.Exit(0);
            }
            else if (secenek == DialogResult.No)
            {
                //Nothing
            }
        }

        private void button_generate_excel_Click(object sender, EventArgs e)
        {
            ASAP_Project.UserPanel.GenerateExcel();
        }

        private void button_create_report_Click(object sender, EventArgs e)
        {
            ASAP_Project.UserPanel.CreateReport();
        }
    }
}