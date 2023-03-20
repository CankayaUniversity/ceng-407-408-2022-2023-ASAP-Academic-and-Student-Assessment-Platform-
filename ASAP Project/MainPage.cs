using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System;

using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.Logging;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

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
            // MAIN PAGE UI PARAMETERS

            panel_adminpanel.Visible = false;
            panel_userpanel.Visible = false;
            panel1.BackColor = Color.Transparent;
            panel2.BackColor = Color.Transparent;
            panel3.BackColor = Color.Transparent;
            panel4.BackColor = Color.Transparent;

            panel5.BackColor = Color.FromArgb(60, Color.Black);
            panel_userpanel.BackColor = Color.FromArgb(60, Color.Black);
            panel_adminpanel.BackColor = Color.FromArgb(60, Color.Black);
            panel_generatexcel.BackColor = Color.FromArgb(60, Color.Black);

            panel5.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, panel5.Width,
            panel5.Height, 30, 30));

            panel_userpanel.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, panel_userpanel.Width,
            panel_userpanel.Height, 30, 30));

            panel_adminpanel.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, panel_adminpanel.Width,
            panel_adminpanel.Height, 30, 30));

            panel_generatexcel.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, panel_generatexcel.Width,
            panel_generatexcel.Height, 30, 30));

            pictureBox1.BackColor = Color.Transparent;
            label_user.Text += LoginScreen.user_name;
            label_user.BackColor = Color.Transparent;
            label_studentcount.BackColor = Color.Transparent;
            label_midtermcount.BackColor = Color.Transparent;
            label_homeworkcount.BackColor = Color.Transparent;
            label_labcount.BackColor = Color.Transparent;
            label_quizcount.BackColor = Color.Transparent;
            label_projectcount.BackColor = Color.Transparent;
            label_derscikticount.BackColor = Color.Transparent;
            label_iscatalog.BackColor = Color.Transparent;
            label_havefinal.BackColor = Color.Transparent;

            button1.BackColor = Color.FromArgb(70, Color.Black);
            button1.FlatAppearance.MouseOverBackColor = Color.FromArgb(90, Color.Black);
            button1.FlatAppearance.MouseDownBackColor = Color.FromArgb(110, Color.Black);
            //button1.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, button1.Width,
            //button1.Height, 5, 5));
            button1.ForeColor = Color.LightGray;

            button_testdrive.BackColor = Color.FromArgb(70, Color.Black);
            button_testdrive.FlatAppearance.MouseOverBackColor = Color.FromArgb(90, Color.Black);
            button_testdrive.FlatAppearance.MouseDownBackColor = Color.FromArgb(110, Color.Black);
            //button_testdrive.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, button_testdrive.Width,
            //button_testdrive.Height, 5, 5));
            button_testdrive.ForeColor = Color.LightGray;

            button_userpanel.BackColor = Color.FromArgb(70, Color.Black);
            button_userpanel.FlatAppearance.MouseOverBackColor = Color.FromArgb(90, Color.Black);
            button_userpanel.FlatAppearance.MouseDownBackColor = Color.FromArgb(110, Color.Black);
            //button_userpanel.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, button_userpanel.Width,
            //button_userpanel.Height, 5, 5));
            button_userpanel.ForeColor = Color.LightGray;

            button_adminpanel.BackColor = Color.FromArgb(70, Color.Black);
            button_adminpanel.FlatAppearance.MouseOverBackColor = Color.FromArgb(90, Color.Black);
            button_adminpanel.FlatAppearance.MouseDownBackColor = Color.FromArgb(110, Color.Black);
            //button_adminpanel.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, button_adminpanel.Width,
            //button_adminpanel.Height, 5, 5));
            button_adminpanel.ForeColor = Color.LightGray;

            button_account.BackColor = Color.FromArgb(70, Color.Black);
            button_account.FlatAppearance.MouseOverBackColor = Color.FromArgb(90, Color.Black);
            button_account.FlatAppearance.MouseDownBackColor = Color.FromArgb(110, Color.Black);
            //button_account.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, button_account.Width,
            //button_account.Height, 5, 5));
            button_account.ForeColor = Color.LightGray;

            button_exit.BackColor = Color.FromArgb(70, Color.Black);
            button_exit.FlatAppearance.MouseOverBackColor = Color.FromArgb(90, Color.Black);
            button_exit.FlatAppearance.MouseDownBackColor = Color.FromArgb(110, Color.Black);
            //button_exit.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, button_exit.Width,
            //button_exit.Height, 5, 5));
            button_exit.ForeColor = Color.LightGray;

            button_create_report.BackColor = Color.FromArgb(70, Color.Black);
            button_create_report.FlatAppearance.MouseOverBackColor = Color.FromArgb(90, Color.Black);
            button_create_report.FlatAppearance.MouseDownBackColor = Color.FromArgb(110, Color.Black);
            //button_userpanel.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, button_userpanel.Width,
            //button_userpanel.Height, 5, 5));
            button_create_report.ForeColor = Color.LightGray;

            button_generate_excel.BackColor = Color.FromArgb(70, Color.Black);
            button_generate_excel.FlatAppearance.MouseOverBackColor = Color.FromArgb(90, Color.Black);
            button_generate_excel.FlatAppearance.MouseDownBackColor = Color.FromArgb(110, Color.Black);
            //button_userpanel.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, button_userpanel.Width,
            //button_userpanel.Height, 5, 5));
            button_generate_excel.ForeColor = Color.LightGray;

            button_edit_report.BackColor = Color.FromArgb(70, Color.Black);
            button_edit_report.FlatAppearance.MouseOverBackColor = Color.FromArgb(90, Color.Black);
            button_edit_report.FlatAppearance.MouseDownBackColor = Color.FromArgb(110, Color.Black);
            //button_userpanel.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, button_userpanel.Width,
            //button_userpanel.Height, 5, 5));
            button_edit_report.ForeColor = Color.LightGray;

            button_transfer_data.BackColor = Color.FromArgb(70, Color.Black);
            button_transfer_data.FlatAppearance.MouseOverBackColor = Color.FromArgb(90, Color.Black);
            button_transfer_data.FlatAppearance.MouseDownBackColor = Color.FromArgb(110, Color.Black);
            //button_userpanel.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, button_userpanel.Width,
            //button_userpanel.Height, 5, 5));
            button_transfer_data.ForeColor = Color.LightGray;




            
        }

        public MainPage()
        {
            InitializeComponent();
            MainPage_Load();
            
        }
                         

        private void button_userpanel_Click(object sender, EventArgs e)
        {
            panel_generatexcel.Visible = false;
            panel_adminpanel.Visible = false;
            panel_adminpanel.Enabled = false;
            panel_userpanel.Enabled = true;
            panel_userpanel.Visible = true;
            panel_userpanel.BringToFront();

        }

        private void button_adminpanel_Click(object sender, EventArgs e)
        {
            panel_generatexcel.Visible = false;
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
            panel_generatexcel.Visible = false;
        }

        private void button_testdrive_Click(object sender, EventArgs e)
        {
            panel_generatexcel.Visible = false;
            try
            {
                ASAP_Project.GoogleDrive.UploadFile();
                // BUG : Dosya seçmeden kapatýnca error veriyor 

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
            panel_generatexcel.Visible = true;
            
        }

        private void button_create_report_Click(object sender, EventArgs e)
        {
            panel_generatexcel.Visible = false;
            ASAP_Project.UserPanel.CreateReport();
        }


        static System.Windows.Forms.Label[] label_mt = new System.Windows.Forms.Label[50];
        static System.Windows.Forms.TextBox[] textbox_mt = new System.Windows.Forms.TextBox[50];
        static System.Windows.Forms.Label[] label_hw = new System.Windows.Forms.Label[50];
        static System.Windows.Forms.TextBox[] textbox_hw = new System.Windows.Forms.TextBox[50];
        static System.Windows.Forms.Label[] label_quiz = new System.Windows.Forms.Label[50];
        static System.Windows.Forms.TextBox[] textbox_quiz = new System.Windows.Forms.TextBox[50];
        static System.Windows.Forms.Label label_final = new System.Windows.Forms.Label();
        static System.Windows.Forms.TextBox textbox_final = new System.Windows.Forms.TextBox();
        static int[] mt_q;
        static int[] hw_q;
        static int[] quiz_q;


        private void textBox_midtermcount_TextChanged(object sender, EventArgs e)
        {

            //261,15

            


            for (int i = 0; i < int.Parse(textBox_midtermcount.Text); i++)
            {
                label_mt[i] = new System.Windows.Forms.Label();
                label_mt[i].Text = "Question Count Midterm " + (i + 1) + " :";
                label_mt[i].ForeColor = Color.White;
                label_mt[i].Size = new System.Drawing.Size(160, 21);
                label_mt[i].Location = new System.Drawing.Point((271) , 15 + (i * 25));
                label_mt[i].Enabled = true;
                label_mt[i].Visible = true;
                label_mt[i].BackColor = Color.Transparent;
                panel_generatexcel.Controls.Add(label_mt[i]);
            }

            

            for (int i = 0; i < int.Parse(textBox_midtermcount.Text); i++)
            {
                textbox_mt[i] = new System.Windows.Forms.TextBox();
                textbox_mt[i].Name = "MidtermQC"+(i + 1);
                textbox_mt[i].Size = new System.Drawing.Size(90, 21);
                textbox_mt[i].Location = new System.Drawing.Point((431), 11 + (i * 25));
                textbox_mt[i].Enabled = true;
                textbox_mt[i].Visible = true;
                panel_generatexcel.Controls.Add(textbox_mt[i]);
            }
        }

        private void button_generatexcel_main_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < int.Parse(textBox_midtermcount.Text); i++)
            {
                mt_q[i] = int.Parse(textbox_mt[i].Text);
            }
            for (int i = 0; i < int.Parse(textBox_homeworkcount.Text); i++)
            {
                hw_q[i] = int.Parse(textbox_hw[i].Text);
            }
            for (int i = 0; i < int.Parse(textBox_quizcount.Text); i++)
            {
                quiz_q[i] = int.Parse(textbox_quiz[i].Text);
            }

            ASAP_Project.UserPanel.GenerateExcel(int.Parse(textBox_studentcount.Text), int.Parse(textBox_midtermcount.Text), 
                int.Parse(textBox_homeworkcount.Text), int.Parse(textBox_labcount.Text), int.Parse(textBox_quizcount.Text), int.Parse(textBox_projectcount.Text),
                int.Parse(textBox_derscikticount.Text), checkBox_iscatalog.Checked, checkBox_havefinal.Checked, mt_q, hw_q,
                quiz_q, int.Parse(textbox_final.Text));
            //ASAP_Project.UserPanel.GenerateExcel();
        }

        private void textBox_homeworkcount_TextChanged(object sender, EventArgs e)
        {
            int lastloc = (label_mt[(int.Parse(textBox_midtermcount.Text)) - 1].Location.Y) + 21;

            for (int i = 0; i < int.Parse(textBox_homeworkcount.Text); i++)
            {
                label_hw[i] = new System.Windows.Forms.Label();
                label_hw[i].Text = "Question Count Homework " + (i + 1) + " :";
                label_hw[i].ForeColor = Color.White;
                label_hw[i].Size = new System.Drawing.Size(170, 21);
                label_hw[i].Location = new System.Drawing.Point((261), lastloc + (i * 25));
                label_hw[i].Enabled = true;
                label_hw[i].Visible = true;
                label_hw[i].BackColor = Color.Transparent;
                panel_generatexcel.Controls.Add(label_hw[i]);
            }


            for (int i = 0; i < int.Parse(textBox_homeworkcount.Text); i++)
            {
                textbox_hw[i] = new System.Windows.Forms.TextBox();
                textbox_hw[i].Size = new System.Drawing.Size(90, 21);
                textbox_hw[i].Location = new System.Drawing.Point((431), lastloc + (i * 25));
                textbox_hw[i].Enabled = true;
                textbox_hw[i].Visible = true;
                panel_generatexcel.Controls.Add(textbox_hw[i]);
            }
        }

        private void textBox_quizcount_TextChanged(object sender, EventArgs e)
        {
            int lastloc = (label_hw[(int.Parse(textBox_homeworkcount.Text)) - 1].Location.Y) + 25;

            for (int i = 0; i < int.Parse(textBox_quizcount.Text); i++)
            {
                label_quiz[i] = new System.Windows.Forms.Label();
                label_quiz[i].Text = "Question Count Quiz " + (i + 1) + " :";
                label_quiz[i].ForeColor = Color.White;
                label_quiz[i].Size = new System.Drawing.Size(140, 21);
                label_quiz[i].Location = new System.Drawing.Point((291), lastloc + (i * 25));
                label_quiz[i].Enabled = true;
                label_quiz[i].Visible = true;
                label_quiz[i].BackColor = Color.Transparent;
                panel_generatexcel.Controls.Add(label_quiz[i]);
            }


            for (int i = 0; i < int.Parse(textBox_quizcount.Text); i++)
            {
                textbox_quiz[i] = new System.Windows.Forms.TextBox();
                textbox_quiz[i].Size = new System.Drawing.Size(90, 21);
                textbox_quiz[i].Location = new System.Drawing.Point((431), lastloc + (i * 25));
                textbox_quiz[i].Enabled = true;
                textbox_quiz[i].Visible = true;
                panel_generatexcel.Controls.Add(textbox_quiz[i]);
            }
        }

        private void checkBox_havefinal_Click(object sender, EventArgs e)
        {
            int lastloc = (label_quiz[(int.Parse(textBox_quizcount.Text)) - 1].Location.Y);

            if (checkBox_havefinal.Checked)
            {
                label_final = new System.Windows.Forms.Label();
                label_final.Text = "Question Count Final " + " :";
                label_final.ForeColor = Color.White;
                label_final.Size = new System.Drawing.Size(140, 21);
                label_final.Location = new System.Drawing.Point((291), lastloc + (25));
                label_final.Enabled = true;
                label_final.Visible = true;
                label_final.BackColor = Color.Transparent;
                panel_generatexcel.Controls.Add(label_final);
                textbox_final = new System.Windows.Forms.TextBox();
                textbox_final.Size = new System.Drawing.Size(90, 21);
                textbox_final.Location = new System.Drawing.Point((431), lastloc + (25));
                textbox_final.Enabled = true;
                textbox_final.Visible = true;
                panel_generatexcel.Controls.Add(textbox_final);
            }
        }
    }
}