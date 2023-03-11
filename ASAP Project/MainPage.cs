namespace ASAP_Project
{
    public partial class MainPage : Form
    {
        public MainPage()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ASAP_Project.GoogleDrive.UploadFile();
        }
    }
}