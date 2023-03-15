namespace ASAP_Project
{
    partial class MainPage
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainPage));
            this.button_userpanel = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel3 = new System.Windows.Forms.Panel();
            this.panel4 = new System.Windows.Forms.Panel();
            this.panel5 = new System.Windows.Forms.Panel();
            this.button1 = new System.Windows.Forms.Button();
            this.button_testdrive = new System.Windows.Forms.Button();
            this.button_exit = new System.Windows.Forms.Button();
            this.button_account = new System.Windows.Forms.Button();
            this.button_adminpanel = new System.Windows.Forms.Button();
            this.panel_userpanel = new System.Windows.Forms.Panel();
            this.button_transferdata = new System.Windows.Forms.Button();
            this.button_transfer_data = new System.Windows.Forms.Button();
            this.button_edit_report = new System.Windows.Forms.Button();
            this.button_create_report = new System.Windows.Forms.Button();
            this.button_generate_excel = new System.Windows.Forms.Button();
            this.panel_adminpanel = new System.Windows.Forms.Panel();
            this.button_deletecourse = new System.Windows.Forms.Button();
            this.button_addcourse = new System.Windows.Forms.Button();
            this.button_updatecourse = new System.Windows.Forms.Button();
            this.panel5.SuspendLayout();
            this.panel_userpanel.SuspendLayout();
            this.panel_adminpanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // button_userpanel
            // 
            this.button_userpanel.Location = new System.Drawing.Point(3, 6);
            this.button_userpanel.Name = "button_userpanel";
            this.button_userpanel.Size = new System.Drawing.Size(184, 62);
            this.button_userpanel.TabIndex = 0;
            this.button_userpanel.Text = "User Panel";
            this.button_userpanel.UseVisualStyleBackColor = true;
            this.button_userpanel.Click += new System.EventHandler(this.button_userpanel_Click);
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.ButtonShadow;
            this.panel1.Dock = System.Windows.Forms.DockStyle.Right;
            this.panel1.Location = new System.Drawing.Point(1214, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(50, 681);
            this.panel1.TabIndex = 1;
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.SystemColors.ButtonShadow;
            this.panel2.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(50, 681);
            this.panel2.TabIndex = 2;
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.SystemColors.ButtonShadow;
            this.panel3.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel3.Location = new System.Drawing.Point(50, 0);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(1164, 50);
            this.panel3.TabIndex = 3;
            // 
            // panel4
            // 
            this.panel4.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.panel4.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel4.Location = new System.Drawing.Point(50, 631);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(1164, 50);
            this.panel4.TabIndex = 4;
            // 
            // panel5
            // 
            this.panel5.BackColor = System.Drawing.SystemColors.ControlDarkDark;
            this.panel5.Controls.Add(this.button1);
            this.panel5.Controls.Add(this.button_testdrive);
            this.panel5.Controls.Add(this.button_exit);
            this.panel5.Controls.Add(this.button_account);
            this.panel5.Controls.Add(this.button_adminpanel);
            this.panel5.Controls.Add(this.button_userpanel);
            this.panel5.Location = new System.Drawing.Point(56, 56);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(193, 569);
            this.panel5.TabIndex = 5;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(76, 333);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(111, 23);
            this.button1.TabIndex = 3;
            this.button1.Text = "TEST CODE";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // button_testdrive
            // 
            this.button_testdrive.Location = new System.Drawing.Point(3, 304);
            this.button_testdrive.Name = "button_testdrive";
            this.button_testdrive.Size = new System.Drawing.Size(184, 23);
            this.button_testdrive.TabIndex = 6;
            this.button_testdrive.Text = "Test Google Drive";
            this.button_testdrive.UseVisualStyleBackColor = true;
            this.button_testdrive.Click += new System.EventHandler(this.button_testdrive_Click);
            // 
            // button_exit
            // 
            this.button_exit.Location = new System.Drawing.Point(6, 535);
            this.button_exit.Name = "button_exit";
            this.button_exit.Size = new System.Drawing.Size(184, 31);
            this.button_exit.TabIndex = 3;
            this.button_exit.Text = "Exit";
            this.button_exit.UseVisualStyleBackColor = true;
            // 
            // button_account
            // 
            this.button_account.Location = new System.Drawing.Point(3, 239);
            this.button_account.Name = "button_account";
            this.button_account.Size = new System.Drawing.Size(184, 59);
            this.button_account.TabIndex = 2;
            this.button_account.Text = "Account";
            this.button_account.UseVisualStyleBackColor = true;
            this.button_account.Click += new System.EventHandler(this.button_account_Click);
            // 
            // button_adminpanel
            // 
            this.button_adminpanel.Location = new System.Drawing.Point(3, 74);
            this.button_adminpanel.Name = "button_adminpanel";
            this.button_adminpanel.Size = new System.Drawing.Size(184, 56);
            this.button_adminpanel.TabIndex = 1;
            this.button_adminpanel.Text = "Admin Panel";
            this.button_adminpanel.UseVisualStyleBackColor = true;
            this.button_adminpanel.Click += new System.EventHandler(this.button_adminpanel_Click);
            // 
            // panel_userpanel
            // 
            this.panel_userpanel.BackColor = System.Drawing.SystemColors.ControlDarkDark;
            this.panel_userpanel.Controls.Add(this.button_transferdata);
            this.panel_userpanel.Controls.Add(this.button_transfer_data);
            this.panel_userpanel.Controls.Add(this.button_edit_report);
            this.panel_userpanel.Controls.Add(this.button_create_report);
            this.panel_userpanel.Controls.Add(this.button_generate_excel);
            this.panel_userpanel.Location = new System.Drawing.Point(255, 56);
            this.panel_userpanel.Name = "panel_userpanel";
            this.panel_userpanel.Size = new System.Drawing.Size(245, 569);
            this.panel_userpanel.TabIndex = 6;
            // 
            // button_transferdata
            // 
            this.button_transferdata.Location = new System.Drawing.Point(3, 295);
            this.button_transferdata.Name = "button_transferdata";
            this.button_transferdata.Size = new System.Drawing.Size(239, 61);
            this.button_transferdata.TabIndex = 4;
            this.button_transferdata.Text = "Transfer Data";
            this.button_transferdata.UseVisualStyleBackColor = true;
            // 
            // button_transfer_data
            // 
            this.button_transfer_data.Location = new System.Drawing.Point(3, 223);
            this.button_transfer_data.Name = "button_transfer_data";
            this.button_transfer_data.Size = new System.Drawing.Size(239, 66);
            this.button_transfer_data.TabIndex = 3;
            this.button_transfer_data.Text = "Transfer Data";
            this.button_transfer_data.UseVisualStyleBackColor = true;
            // 
            // button_edit_report
            // 
            this.button_edit_report.Location = new System.Drawing.Point(3, 150);
            this.button_edit_report.Name = "button_edit_report";
            this.button_edit_report.Size = new System.Drawing.Size(239, 67);
            this.button_edit_report.TabIndex = 2;
            this.button_edit_report.Text = "Edit Report";
            this.button_edit_report.UseVisualStyleBackColor = true;
            // 
            // button_create_report
            // 
            this.button_create_report.Location = new System.Drawing.Point(3, 78);
            this.button_create_report.Name = "button_create_report";
            this.button_create_report.Size = new System.Drawing.Size(239, 66);
            this.button_create_report.TabIndex = 1;
            this.button_create_report.Text = "Create Report";
            this.button_create_report.UseVisualStyleBackColor = true;
            // 
            // button_generate_excel
            // 
            this.button_generate_excel.Location = new System.Drawing.Point(3, 3);
            this.button_generate_excel.Name = "button_generate_excel";
            this.button_generate_excel.Size = new System.Drawing.Size(239, 69);
            this.button_generate_excel.TabIndex = 0;
            this.button_generate_excel.Text = "Generate Excel";
            this.button_generate_excel.UseVisualStyleBackColor = true;
            // 
            // panel_adminpanel
            // 
            this.panel_adminpanel.BackColor = System.Drawing.SystemColors.ControlDarkDark;
            this.panel_adminpanel.Controls.Add(this.button_deletecourse);
            this.panel_adminpanel.Controls.Add(this.button_addcourse);
            this.panel_adminpanel.Controls.Add(this.button_updatecourse);
            this.panel_adminpanel.Location = new System.Drawing.Point(255, 56);
            this.panel_adminpanel.Name = "panel_adminpanel";
            this.panel_adminpanel.Size = new System.Drawing.Size(245, 569);
            this.panel_adminpanel.TabIndex = 7;
            // 
            // button_deletecourse
            // 
            this.button_deletecourse.Location = new System.Drawing.Point(3, 150);
            this.button_deletecourse.Name = "button_deletecourse";
            this.button_deletecourse.Size = new System.Drawing.Size(239, 67);
            this.button_deletecourse.TabIndex = 2;
            this.button_deletecourse.Text = "Delete Course";
            this.button_deletecourse.UseVisualStyleBackColor = true;
            // 
            // button_addcourse
            // 
            this.button_addcourse.Location = new System.Drawing.Point(3, 6);
            this.button_addcourse.Name = "button_addcourse";
            this.button_addcourse.Size = new System.Drawing.Size(239, 69);
            this.button_addcourse.TabIndex = 0;
            this.button_addcourse.Text = "Add Course";
            this.button_addcourse.UseVisualStyleBackColor = true;
            // 
            // button_updatecourse
            // 
            this.button_updatecourse.Location = new System.Drawing.Point(3, 78);
            this.button_updatecourse.Name = "button_updatecourse";
            this.button_updatecourse.Size = new System.Drawing.Size(239, 66);
            this.button_updatecourse.TabIndex = 1;
            this.button_updatecourse.Text = "Update Course";
            this.button_updatecourse.UseVisualStyleBackColor = true;
            // 
            // MainPage
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.ClientSize = new System.Drawing.Size(1264, 681);
            this.Controls.Add(this.panel_userpanel);
            this.Controls.Add(this.panel_adminpanel);
            this.Controls.Add(this.panel5);
            this.Controls.Add(this.panel4);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Name = "MainPage";
            this.Text = "ASAP (Academic and Student Assessment Platform)";
            this.panel5.ResumeLayout(false);
            this.panel_userpanel.ResumeLayout(false);
            this.panel_adminpanel.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private Button button_userpanel;
        private Panel panel1;
        private Panel panel2;
        private Panel panel3;
        private Panel panel4;
        private Panel panel5;
        private Button button_adminpanel;
        private Button button_exit;
        private Button button_account;
        private Panel panel_userpanel;
        private Button button_transferdata;
        private Button button_transfer_data;
        private Button button_edit_report;
        private Button button_create_report;
        private Button button_generate_excel;
        private Panel panel_adminpanel;
        private Button button_deletecourse;
        private Button button_updatecourse;
        private Button button_addcourse;
        private Button button_testdrive;
        private Button button1;
    }
}