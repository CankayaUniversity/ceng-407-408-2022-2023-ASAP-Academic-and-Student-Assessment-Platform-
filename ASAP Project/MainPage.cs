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

        public MainPage()
        {
            InitializeComponent();
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

        private void button1_Click(object sender, EventArgs e)
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
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Sisteminizde Excel kurulu deðil...");
                return;
            }

            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = "Sýra NO";
            xlWorkSheet.Cells[1, 2] = "Ýsim";
            xlWorkSheet.Cells[2, 1] = "1";
            xlWorkSheet.Cells[2, 2] = "Esat";
            xlWorkSheet.Cells[3, 1] = "2";
            xlWorkSheet.Cells[3, 2] = "Emre";

            xlWorkBook.SaveAs("deneme_dosya.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            MessageBox.Show("Excel dosyasý c:\\deneme-dosya.xls adresinde oluþturuldu...");



            //var excelApp = new Excel.Application();
            //excelApp.Workbooks.Add();

            //var xlSheets = excelApp.Sheets as Excel.Sheets;
            //var xlNewSheet = (Excel.Worksheet)xlSheets.Add(xlSheets[1], Type.Missing, Type.Missing, Type.Missing);

            //excelApp.Visible = true;
        }

        private void button_create_report_Click(object sender, EventArgs e)
        {
            //Added code version 1.0 for report genreation, 
            //experiences runtime error in Program.cs, need checking

            //We create two instances for an Excel and a Word File
            Excel.Application excelApp = new Excel.Application();
            Word.Application wordApp = new Word.Application();

            //We pick our Excel file from Pc (Emre's code)
            string oSelectedFile = "";
            System.Windows.Forms.OpenFileDialog oDlg = new System.Windows.Forms.OpenFileDialog();
            if (System.Windows.Forms.DialogResult.OK == oDlg.ShowDialog())
            {
                oSelectedFile = oDlg.FileName;

            }

            //We match our excel 
            Excel.Workbook workbook = excelApp.Workbooks.Open(oSelectedFile);
            Excel.Worksheet worksheet = workbook.ActiveSheet;

            //We create a new Word document
            Word.Document document = wordApp.Documents.Add();

            //We scan our Excel data and add it to our newly
            //created Word document
            for (int i = 1; i <= worksheet.UsedRange.Rows.Count; i++)
            {
                for (int j = 1; j <= worksheet.UsedRange.Columns.Count; j++)
                {
                    string cellValue = worksheet.Cells[i, j].Value.ToString();
                    Word.Range range = document.Content;
                    range.InsertAfter(cellValue + "\t");
                }
                Word.Range rowRange = document.Content;
                rowRange.InsertAfter("\n");
            }
            
            //Save the Word document
            document.SaveAs("/report.docx");

            //Close Excel and Word documents
            workbook.Close();
            excelApp.Quit();

            document.Close();
            wordApp.Quit();
        }
    }
}