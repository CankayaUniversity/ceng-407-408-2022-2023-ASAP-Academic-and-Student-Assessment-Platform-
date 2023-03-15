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
            TreeNode treeNode1 = new TreeNode("Add Course");
            TreeNode treeNode2 = new TreeNode("Update Course");
            TreeNode treeNode3 = new TreeNode("Delete Course");
            TreeNode treeNode4 = new TreeNode("Create Excel");
            TreeNode treeNode5 = new TreeNode("Create Report");
            TreeNode treeNode6 = new TreeNode("Edit Report");
            TreeNode treeNode7 = new TreeNode("Transfer Data");
            TreeNode treeNode8 = new TreeNode("Review Course");
            button_userpanel = new Button();
            panel1 = new Panel();
            panel2 = new Panel();
            panel3 = new Panel();
            panel4 = new Panel();
            panel5 = new Panel();
            button_testdrive = new Button();
            treeView_adminpanel = new TreeView();
            treeView_userpanel = new TreeView();
            button_exit = new Button();
            button_account = new Button();
            button_adminpanel = new Button();
            panel_userpanel = new Panel();
            button_transferdata = new Button();
            button_transfer_data = new Button();
            button_edit_report = new Button();
            button_create_report = new Button();
            button_generate_excel = new Button();
            panel_adminpanel = new Panel();
            button1 = new Button();
            button_deletecourse = new Button();
            button_updatecourse = new Button();
            button_addcourse = new Button();
            panel5.SuspendLayout();
            panel_userpanel.SuspendLayout();
            panel_adminpanel.SuspendLayout();
            SuspendLayout();
            // 
            // button_userpanel
            // 
            button_userpanel.Location = new Point(3, 6);
            button_userpanel.Name = "button_userpanel";
            button_userpanel.Size = new Size(184, 62);
            button_userpanel.TabIndex = 0;
            button_userpanel.Text = "User Panel";
            button_userpanel.UseVisualStyleBackColor = true;
            button_userpanel.Click += button_userpanel_Click;
            // 
            // panel1
            // 
            panel1.BackColor = SystemColors.ButtonShadow;
            panel1.Dock = DockStyle.Right;
            panel1.Location = new Point(786, 0);
            panel1.Name = "panel1";
            panel1.Size = new Size(50, 531);
            panel1.TabIndex = 1;
            // 
            // panel2
            // 
            panel2.BackColor = SystemColors.ButtonShadow;
            panel2.Dock = DockStyle.Left;
            panel2.Location = new Point(0, 0);
            panel2.Name = "panel2";
            panel2.Size = new Size(50, 531);
            panel2.TabIndex = 2;
            // 
            // panel3
            // 
            panel3.BackColor = SystemColors.ButtonShadow;
            panel3.Dock = DockStyle.Top;
            panel3.Location = new Point(50, 0);
            panel3.Name = "panel3";
            panel3.Size = new Size(736, 50);
            panel3.TabIndex = 3;
            // 
            // panel4
            // 
            panel4.BackColor = SystemColors.AppWorkspace;
            panel4.Dock = DockStyle.Bottom;
            panel4.Location = new Point(50, 481);
            panel4.Name = "panel4";
            panel4.Size = new Size(736, 50);
            panel4.TabIndex = 4;
            // 
            // panel5
            // 
            panel5.BackColor = SystemColors.ControlDarkDark;
            panel5.Controls.Add(button_testdrive);
            panel5.Controls.Add(treeView_adminpanel);
            panel5.Controls.Add(treeView_userpanel);
            panel5.Controls.Add(button_exit);
            panel5.Controls.Add(button_account);
            panel5.Controls.Add(button_adminpanel);
            panel5.Controls.Add(button_userpanel);
            panel5.Location = new Point(56, 56);
            panel5.Name = "panel5";
            panel5.Size = new Size(193, 419);
            panel5.TabIndex = 5;
            // 
            // button_testdrive
            // 
            button_testdrive.Location = new Point(3, 304);
            button_testdrive.Name = "button_testdrive";
            button_testdrive.Size = new Size(184, 23);
            button_testdrive.TabIndex = 6;
            button_testdrive.Text = "Test Google Drive";
            button_testdrive.UseVisualStyleBackColor = true;
            button_testdrive.Click += button_testdrive_Click;
            // 
            // treeView_adminpanel
            // 
            treeView_adminpanel.BackColor = SystemColors.ControlDarkDark;
            treeView_adminpanel.BorderStyle = BorderStyle.None;
            treeView_adminpanel.Enabled = false;
            treeView_adminpanel.Location = new Point(3, 136);
            treeView_adminpanel.Name = "treeView_adminpanel";
            treeNode1.Name = "Düğüm0";
            treeNode1.Text = "Add Course";
            treeNode2.Name = "Düğüm1";
            treeNode2.Text = "Update Course";
            treeNode3.Name = "Düğüm2";
            treeNode3.Text = "Delete Course";
            treeView_adminpanel.Nodes.AddRange(new TreeNode[] { treeNode1, treeNode2, treeNode3 });
            treeView_adminpanel.Size = new Size(121, 97);
            treeView_adminpanel.TabIndex = 5;
            treeView_adminpanel.Visible = false;
            // 
            // treeView_userpanel
            // 
            treeView_userpanel.BackColor = SystemColors.ControlDarkDark;
            treeView_userpanel.BorderStyle = BorderStyle.None;
            treeView_userpanel.Enabled = false;
            treeView_userpanel.Location = new Point(3, 74);
            treeView_userpanel.Name = "treeView_userpanel";
            treeNode4.Name = "Düğüm0";
            treeNode4.Text = "Create Excel";
            treeNode5.Name = "Düğüm1";
            treeNode5.Text = "Create Report";
            treeNode6.Name = "Düğüm2";
            treeNode6.Text = "Edit Report";
            treeNode7.Name = "Düğüm3";
            treeNode7.Text = "Transfer Data";
            treeNode8.Name = "Düğüm4";
            treeNode8.Text = "Review Course";
            treeView_userpanel.Nodes.AddRange(new TreeNode[] { treeNode4, treeNode5, treeNode6, treeNode7, treeNode8 });
            treeView_userpanel.Size = new Size(159, 95);
            treeView_userpanel.TabIndex = 4;
            treeView_userpanel.Visible = false;
            // 
            // button_exit
            // 
            button_exit.Location = new Point(3, 385);
            button_exit.Name = "button_exit";
            button_exit.Size = new Size(184, 31);
            button_exit.TabIndex = 3;
            button_exit.Text = "Exit";
            button_exit.UseVisualStyleBackColor = true;
            // 
            // button_account
            // 
            button_account.Location = new Point(3, 239);
            button_account.Name = "button_account";
            button_account.Size = new Size(184, 59);
            button_account.TabIndex = 2;
            button_account.Text = "Account";
            button_account.UseVisualStyleBackColor = true;
            button_account.Click += button_account_Click;
            // 
            // button_adminpanel
            // 
            button_adminpanel.Location = new Point(3, 74);
            button_adminpanel.Name = "button_adminpanel";
            button_adminpanel.Size = new Size(184, 56);
            button_adminpanel.TabIndex = 1;
            button_adminpanel.Text = "Admin Panel";
            button_adminpanel.UseVisualStyleBackColor = true;
            button_adminpanel.Click += button_adminpanel_Click;
            // 
            // panel_userpanel
            // 
            panel_userpanel.BackColor = SystemColors.ControlDarkDark;
            panel_userpanel.Controls.Add(panel_adminpanel);
            panel_userpanel.Controls.Add(button_transferdata);
            panel_userpanel.Controls.Add(button_transfer_data);
            panel_userpanel.Controls.Add(button_edit_report);
            panel_userpanel.Controls.Add(button_create_report);
            panel_userpanel.Controls.Add(button_generate_excel);
            panel_userpanel.Location = new Point(255, 56);
            panel_userpanel.Name = "panel_userpanel";
            panel_userpanel.Size = new Size(245, 419);
            panel_userpanel.TabIndex = 6;
            panel_userpanel.Visible = false;
            // 
            // button_transferdata
            // 
            button_transferdata.Location = new Point(3, 295);
            button_transferdata.Name = "button_transferdata";
            button_transferdata.Size = new Size(239, 61);
            button_transferdata.TabIndex = 4;
            button_transferdata.Text = "Transfer Data";
            button_transferdata.UseVisualStyleBackColor = true;
            // 
            // button_transfer_data
            // 
            button_transfer_data.Location = new Point(3, 223);
            button_transfer_data.Name = "button_transfer_data";
            button_transfer_data.Size = new Size(239, 66);
            button_transfer_data.TabIndex = 3;
            button_transfer_data.Text = "Transfer Data";
            button_transfer_data.UseVisualStyleBackColor = true;
            // 
            // button_edit_report
            // 
            button_edit_report.Location = new Point(3, 150);
            button_edit_report.Name = "button_edit_report";
            button_edit_report.Size = new Size(239, 67);
            button_edit_report.TabIndex = 2;
            button_edit_report.Text = "Edit Report";
            button_edit_report.UseVisualStyleBackColor = true;
            // 
            // button_create_report
            // 
            button_create_report.Location = new Point(3, 78);
            button_create_report.Name = "button_create_report";
            button_create_report.Size = new Size(239, 66);
            button_create_report.TabIndex = 1;
            button_create_report.Text = "Create Report";
            button_create_report.UseVisualStyleBackColor = true;
            // 
            // button_generate_excel
            // 
            button_generate_excel.Location = new Point(3, 3);
            button_generate_excel.Name = "button_generate_excel";
            button_generate_excel.Size = new Size(239, 69);
            button_generate_excel.TabIndex = 0;
            button_generate_excel.Text = "Generate Excel";
            button_generate_excel.UseVisualStyleBackColor = true;
            button_generate_excel.Click += button_generate_excel_Click;
            // 
            // panel_adminpanel
            // 
            panel_adminpanel.BackColor = SystemColors.ControlDarkDark;
            panel_adminpanel.Controls.Add(button1);
            panel_adminpanel.Controls.Add(button_deletecourse);
            panel_adminpanel.Controls.Add(button_addcourse);
            panel_adminpanel.Controls.Add(button_updatecourse);
            panel_adminpanel.Location = new Point(3, 0);
            panel_adminpanel.Name = "panel_adminpanel";
            panel_adminpanel.Size = new Size(245, 419);
            panel_adminpanel.TabIndex = 7;
            panel_adminpanel.Visible = false;
            // 
            // button1
            // 
            button1.Location = new Point(31, 271);
            button1.Name = "button1";
            button1.Size = new Size(111, 23);
            button1.TabIndex = 3;
            button1.Text = "TEST CODE";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // button_deletecourse
            // 
            button_deletecourse.Location = new Point(3, 150);
            button_deletecourse.Name = "button_deletecourse";
            button_deletecourse.Size = new Size(239, 67);
            button_deletecourse.TabIndex = 2;
            button_deletecourse.Text = "Delete Course";
            button_deletecourse.UseVisualStyleBackColor = true;
            // 
            // button_updatecourse
            // 
            button_updatecourse.Location = new Point(3, 78);
            button_updatecourse.Name = "button_updatecourse";
            button_updatecourse.Size = new Size(239, 66);
            button_updatecourse.TabIndex = 1;
            button_updatecourse.Text = "Update Course";
            button_updatecourse.UseVisualStyleBackColor = true;
            // 
            // button_addcourse
            // 
            button_addcourse.Location = new Point(3, 6);
            button_addcourse.Name = "button_addcourse";
            button_addcourse.Size = new Size(239, 69);
            button_addcourse.TabIndex = 0;
            button_addcourse.Text = "Add Course";
            button_addcourse.UseVisualStyleBackColor = true;
            // 
            // MainPage
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(836, 531);
            Controls.Add(panel_userpanel);
            Controls.Add(panel5);
            Controls.Add(panel4);
            Controls.Add(panel3);
            Controls.Add(panel2);
            Controls.Add(panel1);
            Name = "MainPage";
            Text = "ASAP (Academic and Student Assessment Platform)";
            panel5.ResumeLayout(false);
            panel_userpanel.ResumeLayout(false);
            panel_adminpanel.ResumeLayout(false);
            ResumeLayout(false);
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
        private TreeView treeView_userpanel;
        private TreeView treeView_adminpanel;
        private Button button_testdrive;
        private Button button1;
    }
}