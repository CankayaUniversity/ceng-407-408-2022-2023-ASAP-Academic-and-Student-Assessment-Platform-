namespace ASAP_Project
{
    public partial class MainPage : Form
    {
        public MainPage()
        {         
            InitializeComponent();
            panel_adminpanel.Visible = false;
            panel_userpanel.Visible = false;
        }

        private void button_userpanel_Click(object sender, EventArgs e)
        {
            panel_adminpanel.Visible = false;
            panel_adminpanel.Enabled = false;
            panel_userpanel.Enabled = true;
            panel_userpanel.Visible = true;
            panel_userpanel.BringToFront();
            treeView_userpanel.Enabled = true;
            treeView_userpanel.Visible = true;
            button_adminpanel.Location = new Point(3, 175);
            treeView_adminpanel.Enabled = false;
            treeView_adminpanel.Visible = false;

        }

        private void button_adminpanel_Click(object sender, EventArgs e)
        {
            panel_userpanel.Visible=false;
            panel_userpanel.Enabled = false;
            panel_adminpanel.Enabled = true;
            panel_adminpanel.Visible=true;
            panel_adminpanel.BringToFront();
            button_adminpanel.Location = new Point(3, 74);
            treeView_adminpanel.Enabled = true;
            treeView_adminpanel.Visible = true;
            treeView_userpanel.Enabled = false;
            treeView_userpanel.Visible = false;
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
    }
}