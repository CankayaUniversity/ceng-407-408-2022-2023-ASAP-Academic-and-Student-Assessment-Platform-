using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System;

using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
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
            //button_adminpanel.Location = System.Drawing.Point(3,175);
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
            //button_adminpanel.Location = Point(3,74);
            treeView_adminpanel.Enabled = true;
            treeView_adminpanel.Visible = true;
            treeView_userpanel.Enabled = false;
            treeView_userpanel.Visible = false;
        }

        private void button_account_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Sisteminizde Excel kurulu de�il...");
                return;
            }

            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = "S�ra NO";
            xlWorkSheet.Cells[1, 2] = "�sim";
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

            MessageBox.Show("Excel dosyas� c:\\deneme-dosya.xls adresinde olu�turuldu...");
        }

        private void button_testdrive_Click(object sender, EventArgs e)
        {
            try
            {
                // ASAP_Project.GoogleDrive.UploadFile();

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
    }
}