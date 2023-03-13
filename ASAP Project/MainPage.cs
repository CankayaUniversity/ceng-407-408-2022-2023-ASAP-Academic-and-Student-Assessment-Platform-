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
        }

        private void button_adminpanel_Click(object sender, EventArgs e)
        {
            panel_userpanel.Visible=false;
            panel_userpanel.Enabled = false;
            panel_adminpanel.Enabled = true;
            panel_adminpanel.Visible=true;
            panel_adminpanel.BringToFront();
        }
    }
}