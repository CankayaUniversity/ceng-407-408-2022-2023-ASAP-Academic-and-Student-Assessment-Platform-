﻿using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data.OleDb;
using Microsoft.Win32;
using System.IO;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Range = Microsoft.Office.Interop.Excel.Range;
using System.ComponentModel;
using System.Collections.ObjectModel;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Microsoft.Office.Interop.Word;
using System.Threading;

namespace ASAP_Project
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
            
        }

        private void button_userpanel_Click(object sender, RoutedEventArgs e)
        {
            grid_adminpanel.Visibility = Visibility.Hidden;
            grid_addcourse.Visibility = Visibility.Collapsed;
            grid_changepassword.Visibility = Visibility.Collapsed;
            grid_accountpanel.Visibility = Visibility.Collapsed;

            grid_userpanel.Visibility = Visibility.Visible;
        }

        private void button_exit_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void button_adminpanel_Click(object sender, RoutedEventArgs e)
        {
            if (UserData.role == "admin")
            {
                grid_userpanel.Visibility = Visibility.Collapsed;
                grid_accountpanel.Visibility = Visibility.Collapsed;
                grid_generate_excel.Visibility = Visibility.Collapsed;
                grid_createreport.Visibility = Visibility.Collapsed;
                grid_reviewcourse.Visibility = Visibility.Collapsed;
                grid_changepassword.Visibility = Visibility.Collapsed;

                grid_adminpanel.Visibility = Visibility.Visible;
            }
            else
            {
                MessageBox.Show("You have not a permission for that");
            }
            //
        }


        
        private void button_generate_excel_Click(object sender, RoutedEventArgs e)
        {
            grid_createreport.Visibility = Visibility.Collapsed;
            grid_reviewcourse.Visibility = Visibility.Collapsed;
            grid_changepassword.Visibility = Visibility.Collapsed;

            combobox_courselist.Items.Clear();

            if (grid_generate_excel.Visibility == Visibility.Visible)
            {
                grid_generate_excel.Visibility = Visibility.Hidden;
            }
            else
            {
                grid_generate_excel.Visibility = Visibility.Visible;

                List<string> list = new List<string>();
                list = GoogleDrive.getCourseList();

                foreach (var item in list)
                {
                    combobox_courselist.Items.Add(item);
                }

            }

        }

        System.Windows.Controls.TextBox[] midtermtextbox = new System.Windows.Controls.TextBox[10];

        //Textboxtaki verilerin değişikliğini kontrol etmek için

        private List<System.Windows.Controls.Label> midtermlabelList = new List<System.Windows.Controls.Label>();
        private List<System.Windows.Controls.TextBox> midtermtextboxList = new List<System.Windows.Controls.TextBox>();

        private bool programmaticChange = false;

        private void textbox_midtermcount_TextChanged(object sender, TextChangedEventArgs e)
        {

            if (string.IsNullOrWhiteSpace(textbox_midtermcount.Text))
            {
                return; 
            }

            
            foreach (var label in midtermlabelList)
            {
                grid_generate_excel.Children.Remove(label);
            }
            foreach (var textbox in midtermtextboxList)
            {
                grid_generate_excel.Children.Remove(textbox);
            }
            foreach (var label in homeworklabelList)
            {
                grid_generate_excel.Children.Remove(label);
            }
            foreach (var textbox in homeworktextboxList)
            {
                grid_generate_excel.Children.Remove(textbox);
            }

            midtermlabelList.Clear();
            midtermtextboxList.Clear();
            homeworklabelList.Clear();
            homeworktextboxList.Clear();

            int midtermCount = int.Parse(textbox_midtermcount.Text);

            for (int i = 0; i < midtermCount; i++)
            {
                System.Windows.Controls.Label midtermlabel = new System.Windows.Controls.Label();
                midtermlabel.Name = "qCountmt" + (i + 1);
                midtermlabel.HorizontalAlignment = HorizontalAlignment.Left;
                midtermlabel.VerticalAlignment = VerticalAlignment.Top;
                midtermlabel.Width = 160;
                midtermlabel.Height = 30;
                midtermlabel.Opacity = 0.8;
                midtermlabel.Content = "Question Count Midterm " + (i + 1) + " :";
                midtermlabel.Margin = new Thickness(272, (i * 30) + 29, 0, 0);
                midtermlabel.Foreground = Brushes.White;
                midtermlabel.Visibility = Visibility.Visible;

                grid_generate_excel.Children.Add(midtermlabel);
                midtermlabelList.Add(midtermlabel); // add to the list

                System.Windows.Controls.TextBox midtermtextbox = new System.Windows.Controls.TextBox();
                midtermtextbox.Name = "qTextboxmidterm" + (i + 1);
                midtermtextbox.HorizontalAlignment = HorizontalAlignment.Left;
                midtermtextbox.VerticalAlignment = VerticalAlignment.Top;
                midtermtextbox.Width = 70;
                midtermtextbox.Height = 15;
                midtermtextbox.Margin = new Thickness(440, (i * 30) + 35, 0, 0);

                grid_generate_excel.Children.Add(midtermtextbox);
                midtermtextboxList.Add(midtermtextbox); // add to the list
            }
            programmaticChange = true;
            textbox_homeworkcount.Text = " ";
            programmaticChange = false;
        }

        System.Windows.Controls.TextBox[] homeworktextbox = new System.Windows.Controls.TextBox[10];

        //for bugs
        private List<System.Windows.Controls.Label> homeworklabelList = new List<System.Windows.Controls.Label>();
        private List<System.Windows.Controls.TextBox> homeworktextboxList = new List<System.Windows.Controls.TextBox>();



        private void textbox_homeworkcount_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textbox_homeworkcount.Text))
            {
                return;  // Exit the method
            }

            // Clear the grid and the lists first
            foreach (var label in homeworklabelList)
            {
                grid_generate_excel.Children.Remove(label);
            }
            foreach (var textbox in homeworktextboxList)
            {
                grid_generate_excel.Children.Remove(textbox);
            }

            checkbox_havefinal.IsChecked = false;
            grid_generate_excel.Children.Remove(final_label);
            grid_generate_excel.Children.Remove(finaltextbox);
            homeworklabelList.Clear();
            homeworktextboxList.Clear();

            int lastps = int.Parse(textbox_midtermcount.Text);
            int last = lastps * 30;

            int homeworkCount = int.Parse(textbox_homeworkcount.Text);

            for (int i = 0; i < homeworkCount; i++)
            {
                System.Windows.Controls.Label homeworklabel = new System.Windows.Controls.Label();
                homeworklabel.Name = "qCountmt" + (i + 1);
                homeworklabel.HorizontalAlignment = HorizontalAlignment.Left;
                homeworklabel.VerticalAlignment = VerticalAlignment.Top;
                homeworklabel.Width = 170;
                homeworklabel.Height = 30;
                homeworklabel.Opacity = 0.8;
                homeworklabel.Content = "Question Count Homework " + (i + 1) + " :";
                homeworklabel.Margin = new Thickness(272, (i * 30) + last + 29, 0, 0);
                homeworklabel.Foreground = Brushes.White;
                homeworklabel.Visibility = Visibility.Visible;

                grid_generate_excel.Children.Add(homeworklabel);
                homeworklabelList.Add(homeworklabel); // add to the list

                System.Windows.Controls.TextBox homeworktextbox = new System.Windows.Controls.TextBox();
                homeworktextbox.Name = "qTextboxhomework" + (i + 1);
                homeworktextbox.HorizontalAlignment = HorizontalAlignment.Left;
                homeworktextbox.VerticalAlignment = VerticalAlignment.Top;
                homeworktextbox.Width = 70;
                homeworktextbox.Height = 15;
                homeworktextbox.Margin = new Thickness(440, (i * 30) + last + 35, 0, 0);

                grid_generate_excel.Children.Add(homeworktextbox);
                homeworktextboxList.Add(homeworktextbox); // add to the list
            }
           

        }



        private void button_transferdata_Click(object sender, RoutedEventArgs e)
        {
            grid_generate_excel.Visibility = Visibility.Hidden;
            if (grid_transferdata.Visibility == Visibility.Visible)
            {
                grid_transferdata.Visibility = Visibility.Hidden;
            }
            else
            {
                grid_transferdata.Visibility = Visibility.Visible;
            }
        }

        private void button_transferdatatogoogledrive_Click(object sender, RoutedEventArgs e)
        {
            GoogleDrive.UploadFile();
        }

        private void button_ok_Click(object sender, RoutedEventArgs e)
        {

            var openFileDialog = new OpenFileDialog();

            if (openFileDialog.ShowDialog() == true)
            {
                string filename = openFileDialog.FileName;
                string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename + ";Extended Properties=\"Excel 12.0;HDR=YES\"";
                OleDbConnection connection = new OleDbConnection(connectionString);
                connection.Open();
                OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT * FROM [Course_Grading_Constraints$]", connection);
                System.Data.DataTable dt = new System.Data.DataTable();
                adapter.Fill(dt);
                connection.Close();
                datagrid_reviewcourse.ItemsSource = dt.DefaultView;
            }
        }
        
        private void button_reviewcourse_Click(object sender, RoutedEventArgs e)
        {
            grid_generate_excel.Visibility = Visibility.Collapsed;
            grid_createreport.Visibility = Visibility.Collapsed;
            grid_changepassword.Visibility = Visibility.Collapsed;
            if (grid_reviewcourse.Visibility == Visibility.Visible)
            {
                grid_reviewcourse.Visibility = Visibility.Hidden;
            }
            else
            {
                grid_reviewcourse.Visibility = Visibility.Visible;
            }
        }
        System.Windows.Controls.TextBox finaltextbox = new System.Windows.Controls.TextBox();

        System.Windows.Controls.Label final_label = new System.Windows.Controls.Label();
        

        private void checkbox_havefinal_Checked(object sender, RoutedEventArgs e)
        {
            int lastps = int.Parse(textbox_homeworkcount.Text);
            int last = lastps * 30 + 30;
            final_label = new System.Windows.Controls.Label();
            final_label.Name = "qCountfinal";
            final_label.HorizontalAlignment = HorizontalAlignment.Left;
            final_label.VerticalAlignment = VerticalAlignment.Top;
            final_label.Width = 170;
            final_label.Height = 30;
            final_label.Opacity = 0.8;
            final_label.Content = "Question Count Final " + " :";
            final_label.Margin = new Thickness(272, 30 + last + 29, 0, 0);
            final_label.Foreground = Brushes.White;
            final_label.Visibility = Visibility.Visible;
            grid_generate_excel.Children.Add(final_label);


            

            finaltextbox = new System.Windows.Controls.TextBox();
            finaltextbox.Name = "qTextboxfinal";
            finaltextbox.HorizontalAlignment = HorizontalAlignment.Left;
            finaltextbox.VerticalAlignment = VerticalAlignment.Top;
            finaltextbox.Width = 70;
            finaltextbox.Height = 15;
            finaltextbox.Margin = new Thickness(440, 30 + last + 35, 0, 0);
            grid_generate_excel.Children.Add(finaltextbox);
        }

        //This code takes the excel file
        //finds the student information and take it into a 2d array for us to use later on
        //I said return an object because, we will use this while we send data to generate excel or ExcelCalculator
        //which works with Report Generator(CreateReport)
        public String[,] Name_taker(Excel.Workbook wb, ref int Student_Count)
        {
            String[,] StuInfo = null;
            //Now this will be the code we take from admin page to drive and to this snippet of code
            int totalWorksheets = wb.Worksheets.Count;
            
            for (int i = totalWorksheets; i > 0; i--)
            {
                Excel.Worksheet worksheet = (Excel.Worksheet)wb.Worksheets[i];
                if(worksheet.Name == "Students")
                {
                    Student_Count = Student_counter(worksheet);
                    StuInfo = new String[Student_Count, 3];
                    for (int j = 2; j <Student_Count + 2; j++)
                    {
                        for(int k = 2; k < 3 + 2; k++) //to make it more understandable I wrote 3+2 instead of 5
                        {
                            StuInfo[j - 2, k - 2] = Convert.ToString(worksheet.Cells[j,k].Value);
                        }
                    }
                    break;
                }
            }
            return StuInfo;
        }

        //This one calculates the student_no of that sheet
        public int Student_counter(Excel.Worksheet worksheet)
        {
            int Student_no = 0;
            for (int i = 2; i < worksheet.Cells.Columns.Count; i++)
            {
                if (worksheet.Cells[i, 1].Value == i - 1)
                {
                    Student_no++;
                }
                else if (worksheet.Cells[i, 1].Value == null)
                {
                    break;
                }
            }
            return Student_no;
        }

        //This one is responsible for handling the button click of "Generate Excel"
        private void button_generate_excel_btnr_Click(object sender, RoutedEventArgs e)
        {
            int[] midtermqcount = new int[10];
            for (int i = 0; i < midtermtextboxList.Count; i++)
            {
                if (midtermtextboxList[i] == null || int.Parse(midtermtextboxList[i].Text) == 0)
                {
                    break;
                }
                else
                {
                    midtermqcount[i] = int.Parse(midtermtextboxList[i].Text);
                }
            }
            int[] homeworkqcount = new int[10];
            for (int i = 0; i < homeworktextboxList.Count; i++)
            {
                if (homeworktextboxList[i] == null || int.Parse(homeworktextboxList[i].Text) == 0)
                {
                    break;
                }
                else
                {
                    homeworkqcount[i] = int.Parse(homeworktextboxList[i].Text);
                }
            }

            //Final booleanı girilmeyince hata veriyor
            //Bir değeri diğerinden önnce giremiyorum, örneüin homework sayısı vermeden final durumu işaretlemeyi deneyince hata veriyor.
            //Herhangi bir değeri girip silince hata veriyor
            UserPanel userPanel = new UserPanel();

            MemoryStream secilenexcel = new MemoryStream();
            secilenexcel = GoogleDrive.GetFile(combobox_courselist.SelectedItem.ToString());
            // Converts MemoryStream to byte[]
            byte[] excelData = secilenexcel.ToArray();

            // Saves byte[] as a temporary file
            string tempFilePath = System.IO.Path.GetTempFileName();
            File.WriteAllBytes(tempFilePath, excelData);

            // Opens the temporary file with Excel Interop
            Application excelApp = new Application();
            Workbook wb = excelApp.Workbooks.Open(tempFilePath);
            int Student_Count = 0;
            //This also get editted in the Name_taker function
            //thanks to the referencing (I guess)
            //-Tan :D
            String[,] info = Name_taker(wb, ref Student_Count);
            Excel.Worksheet DCPCworksheet = null;
            foreach (Excel.Worksheet worksheet in wb.Sheets)
            {
                if(worksheet.Name == "DC-PC")
                {
                    DCPCworksheet = worksheet;
                    break;
                }
            }
            userPanel.GenerateExcel(info, Student_Count, int.Parse(textbox_midtermcount.Text), int.Parse(textbox_homeworkcount.Text), int.Parse(textbox_labcount.Text),
                int.Parse(textbox_quizcount.Text), int.Parse(textbox_projectcount.Text), checkbox_havefinal.IsChecked ?? false, midtermqcount, homeworkqcount, int.Parse(finaltextbox.Text),DCPCworksheet);
        }

        private void button_createreport_Click(object sender, RoutedEventArgs e)
        {
            grid_generate_excel.Visibility = Visibility.Collapsed;
            grid_reviewcourse.Visibility = Visibility.Collapsed;
            grid_changepassword.Visibility = Visibility.Collapsed;
            combobox_createreport.Items.Clear();

            if (grid_createreport.Visibility == Visibility.Visible)
            {
                grid_createreport.Visibility = Visibility.Hidden;
            }
            else
            {
                grid_createreport.Visibility = Visibility.Visible;

                List<string> list = new List<string>();
                list = GoogleDrive.getGenExcelList();

                foreach (var item in list)
                {
                    combobox_createreport.Items.Add(item);
                }

            }
            //UserPanel.CreateReport();
        }

        private void button_downloadexcel_Click(object sender, RoutedEventArgs e)
        {
            GoogleDrive.UploadFile();
        }

        private void button_selectexcelfile_Click(object sender, RoutedEventArgs e)
        {
            MemoryStream secilenexcel = new MemoryStream();
            secilenexcel = GoogleDrive.GetFile(combobox_createreport.SelectedItem.ToString());
            // Converts MemoryStream to byte[]
            byte[] excelData = secilenexcel.ToArray();

            // Saves byte[] as a temporary file
            string tempFilePath = System.IO.Path.GetTempFileName();
            File.WriteAllBytes(tempFilePath, excelData);

            // Opens the temporary file with Excel Interop
            Application excelApp = new Application();
            Workbook wb = excelApp.Workbooks.Open(tempFilePath);
            UserPanel userPanel = new UserPanel();
            userPanel.CreateReport(excelApp, wb);
        }

        private void button_addcourse_Click(object sender, RoutedEventArgs e)
        {
            if (grid_addcourse.Visibility == Visibility.Visible)
            {
                grid_addcourse.Visibility = Visibility.Hidden;
            }
            else 
            {
                grid_addcourse.Visibility = Visibility.Visible;
            }
            
        }
        

        public ObservableCollection<DC_PC_CheckBoxTable> Rows { get; set; }

        private void textbox_howmanydc_TextChanged(object sender, TextChangedEventArgs e)
        {
            int dc_count = int.Parse(textbox_howmanydc.Text);

            System.Data.DataTable dc_pc_datatable = new System.Data.DataTable();

            DataContext = this;

            Rows = new ObservableCollection<DC_PC_CheckBoxTable>();

            for (int i = 0; i < dc_count; i++)
            {
                Rows.Add(new DC_PC_CheckBoxTable
                {
                    DCPC = "DÇ " + (i + 1).ToString(),
                    PC1 = false,
                    PC2 = false,
                    PC3 = false,
                    PC4 = false,
                    PC5 = false,
                    PC6 = false,
                    PC7 = false,
                    PC8 = false,
                    PC9 = false,
                    PC10 = false,
                    PC11 = false,
                });
            }

            datagrid_addcourse.ItemsSource = Rows;

            
        }

        public class DC_PC_CheckBoxTable
        {
            public string DCPC { get; set; }           
            public bool PC1 { get; set; }
            public bool PC2 { get; set; }
            public bool PC3 { get; set; }
            public bool PC4 { get; set; }
            public bool PC5 { get; set; }
            public bool PC6 { get; set; }
            public bool PC7 { get; set; }
            public bool PC8 { get; set; }
            public bool PC9 { get; set; }
            public bool PC10 { get; set; }
            public bool PC11 { get; set; }
        }

        private void button_addcourseexportexcel_Click(object sender, RoutedEventArgs e)
        {
            System.Data.DataTable dataTable = new System.Data.DataTable();

            Application excel = new Application();

            Workbook workbook = excel.Workbooks.Add();

            Worksheet worksheet = workbook.ActiveSheet;

            worksheet.Name = "DC-PC";

            workbook.Title = (textbox_coursecode.Text).ToString();

            foreach (DataGridColumn col in datagrid_addcourse.Columns)
            {
                dataTable.Columns.Add(col.Header.ToString(), typeof(string));
            }


            foreach (var row in datagrid_addcourse.Items)
            {
                DC_PC_CheckBoxTable rowView = (DC_PC_CheckBoxTable)row;
                DataRow dataRow = dataTable.NewRow();

                dataRow[0] = rowView.DCPC;

                if (rowView.PC1 == true)
                {
                    dataRow[1] = 1;
                }
                else
                {
                    dataRow[1] = 0;
                }
                if (rowView.PC2 == true)
                {
                    dataRow[2] = 1;
                }
                else
                {
                    dataRow[2] = 0;
                }
                if (rowView.PC3 == true)
                {
                    dataRow[3] = 1;
                }
                else
                {
                    dataRow[3] = 0;
                }
                if (rowView.PC4 == true)
                {
                    dataRow[4] = 1;
                }
                else
                {
                    dataRow[4] = 0;
                }
                if (rowView.PC5 == true)
                {
                    dataRow[5] = 1;
                }
                else
                {
                    dataRow[5] = 0;
                }
                if (rowView.PC6 == true)
                {
                    dataRow[6] = 1;
                }
                else
                {
                    dataRow[6] = 0;
                }
                if (rowView.PC7 == true)
                {
                    dataRow[7] = 1;
                }
                else
                {
                    dataRow[7] = 0;
                }
                if (rowView.PC8 == true)
                {
                    dataRow[8] = 1;
                }
                else
                {
                    dataRow[8] = 0;
                }
                if (rowView.PC9 == true)
                {
                    dataRow[9] = 1;
                }
                else
                {
                    dataRow[9] = 0;
                }
                if (rowView.PC10 == true)
                {
                    dataRow[10] = 1;
                }
                else
                {
                    dataRow[10] = 0;
                }
                if (rowView.PC11 == true)
                {
                    dataRow[11] = 1;
                }
                else
                {
                    dataRow[11] = 0;
                }
                

                dataTable.Rows.Add(dataRow);
            }

            

            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                worksheet.Cells[1, i + 1] = dataTable.Columns[i].ColumnName;
            }
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                for (int j = 0; j < dataTable.Columns.Count; j++)
                {
                    worksheet.Cells[i + 2, j + 1] = dataTable.Rows[i][j].ToString();
                }
            }

            Worksheet student = workbook.Worksheets.Add();
            student.Name = "Students";
            student.Cells[1, 1] = "Id";
            student.Cells[1, 2] = "Student ID";
            student.Cells[1, 3] = "Student Name";
            student.Cells[1, 4] = "Student Surname";

            for (int i = 1; i < int.Parse(textbox_coursestudentcount.Text) + 1; i++)
            {
                student.Cells[i + 1, 1] = i;
            }

            string fileName = textbox_coursecode.Text + "-" + textbox_courseyear.Text;
            string appDataFolder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string filePath = System.IO.Path.Combine(appDataFolder, fileName);

            workbook.SaveAs(filePath);
            MessageBox.Show("Your excel file saved as this location: " + filePath);

            excel.Visible = true;
        }

        private void button_account_Click(object sender, RoutedEventArgs e)
        {

            grid_userpanel.Visibility = Visibility.Collapsed;
            grid_adminpanel.Visibility = Visibility.Collapsed;
            grid_generate_excel.Visibility = Visibility.Collapsed;
            grid_createreport.Visibility = Visibility.Collapsed;
            grid_reviewcourse.Visibility = Visibility.Collapsed;
            grid_addcourse.Visibility = Visibility.Collapsed;
            if (grid_accountpanel.Visibility == Visibility.Visible)
            {
                grid_accountpanel.Visibility = Visibility.Hidden;
            }
            else
            {
                grid_accountpanel.Visibility = Visibility.Visible;
            }
        }

        private void button_ChangePassword_Click(object sender, RoutedEventArgs e)
        {
            if (grid_changepassword.Visibility == Visibility.Visible)
            {
                grid_changepassword.Visibility = Visibility.Hidden;
            }
            else
            {
                grid_changepassword.Visibility = Visibility.Visible;
            }
        }

        private void button_addcourseupload_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            openFileDialog.ShowDialog();

            GoogleDrive.UploadCourse(openFileDialog.FileName);
        }

        private void button_deletecourse_Click(object sender, RoutedEventArgs e)
        {
            combobox_deletecourse.Items.Clear();

            if (grid_deletecourse.Visibility == Visibility.Visible)
            {
                grid_deletecourse.Visibility = Visibility.Hidden;
            }
            else
            {
                grid_deletecourse.Visibility = Visibility.Visible;

                List<string> list = new List<string>();
                list = GoogleDrive.getCourseList();

                foreach (var item in list)
                {
                    combobox_deletecourse.Items.Add(item);
                }

            }
        }

        private void button_deletecoursebtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string filepath = combobox_deletecourse.SelectedItem.ToString();
                GoogleDrive.DeleteFile(filepath);

                MessageBox.Show("File :" + filepath + " deleted from google drive.");
            }
            catch(Exception ex)
            {
                MessageBox.Show($"Error deleting file to Google Drive: {ex.Message}");
            }

            
        }







    private void Update_Password_Click(object sender, RoutedEventArgs e)
    {
            string oldPassword = textbox_oldpassword.Text;
            string newPassword = textbox_newpassword.Text;


            MemoryStream userdata = GoogleDrive.GetFile("UserInfo.xlsx");
            byte[] excelData = userdata.ToArray();
            string tempFilePath = System.IO.Path.GetTempFileName();

            File.WriteAllBytes(tempFilePath, excelData);


            try
            {
                // Excel dosyasını güncelle
                UpdateExcelFile(tempFilePath, oldPassword, newPassword);


                var fileMetadata = new Google.Apis.Drive.v3.Data.File()
                {
                    Name = "UserInfo.xlsx",
                    
                };
                GoogleDrive.DeleteUserInfo("UserInfo.xlsx");
                GoogleDrive.UploadUserInfo(tempFilePath);


                MessageBox.Show("Parola başarıyla güncellendi.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Excel dosyasını indirme/güncelleme hatası: " + ex.Message);
            }
        }

        private void UpdateExcelFile(string tempFilePath, string oldPassword, string newPassword)
        {

            Application excelApp = new Application();
            Workbook wb = excelApp.Workbooks.Open(tempFilePath);
            
            var worksheet = wb.Sheets[1];
            var range = worksheet.UsedRange;
            try
            {
                // foreach (Excel.Worksheet worksheet in wb.Sheets)
                // {
                //     if (worksheet.Name == oldPassword)
                //     {
                //         newPassword = oldPassword;
                //         break;
                //     }
                // }
                for (int row = 2; row <= range.Rows.Count; row++) // İlk satırı başlık kabul ediyoruz, bu yüzden 2'den başlıyoruz
                {
                    string password = Convert.ToString((range.Cells[row, 3] as Excel.Range).Value2);

                    if (password == oldPassword)
                    {
                        // Parola bulundu, güncelle
                        (range.Cells[row, 3] as Excel.Range).Value2 = newPassword;
                        break;
                    }
                }
            }

            catch (Exception ex)
            {
                throw new Exception("Parola güncelleme hatası: " + ex.Message);
            }

            finally
            {
                // Excel işlemlerini temizle ve kaynakları serbest bırak
                wb.Save();
                wb.Close();
                excelApp.Quit();
                ReleaseObject(wb);
                ReleaseObject(excelApp);
            }
        }

        private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("COM nesnesi serbest bırakma hatası: " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    } 
}
