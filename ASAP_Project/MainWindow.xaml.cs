using Microsoft.Office.Interop.Excel;
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
            grid_userpanel.Visibility = Visibility.Visible;
        }

        private void button_exit_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void button_adminpanel_Click(object sender, RoutedEventArgs e)
        {
            grid_userpanel.Visibility = Visibility.Hidden;
            grid_adminpanel.Visibility = Visibility.Visible;
            //
        }

        private void button_generate_excel_Click(object sender, RoutedEventArgs e)
        {
            if (grid_generate_excel.Visibility == Visibility.Visible)
            {
                grid_generate_excel.Visibility = Visibility.Hidden;
            }
            else
            {
                grid_generate_excel.Visibility = Visibility.Visible;
            }

        }

        System.Windows.Controls.TextBox[] midtermtextbox = new System.Windows.Controls.TextBox[10];
        
        private void textbox_midtermcount_TextChanged(object sender, TextChangedEventArgs e)
        {
            

            System.Windows.Controls.Label[] midtermlabel = new System.Windows.Controls.Label[int.Parse(textbox_midtermcount.Text)];
            for (int i = 0; i < int.Parse(textbox_midtermcount.Text); i++)
            {
                midtermlabel[i] = new System.Windows.Controls.Label();
                midtermlabel[i].Name = "qCountmt" + (i + 1);
                midtermlabel[i].HorizontalAlignment = HorizontalAlignment.Left;
                midtermlabel[i].VerticalAlignment = VerticalAlignment.Top;
                midtermlabel[i].Width = 160;
                midtermlabel[i].Height = 30;
                midtermlabel[i].Opacity = 0.8;
                midtermlabel[i].Content = "Question Count Midterm " + (i + 1) + " :";
                midtermlabel[i].Margin = new Thickness(272, (i * 30) + 29, 0, 0);
                midtermlabel[i].Foreground = Brushes.White;
                midtermlabel[i].Visibility = Visibility.Visible;
                grid_generate_excel.Children.Add(midtermlabel[i]);
            }

            for (int i = 0; i < int.Parse(textbox_midtermcount.Text); i++)
            {
                midtermtextbox[i] = new System.Windows.Controls.TextBox();
                midtermtextbox[i].Name = "qTextboxmidterm" + (i + 1);
                midtermtextbox[i].HorizontalAlignment = HorizontalAlignment.Left;
                midtermtextbox[i].VerticalAlignment = VerticalAlignment.Top;
                midtermtextbox[i].Width = 70;
                midtermtextbox[i].Height = 15;
                midtermtextbox[i].Margin = new Thickness(440, (i * 30) + 35, 0, 0);
                grid_generate_excel.Children.Add(midtermtextbox[i]);
            }

        }
        System.Windows.Controls.TextBox[] homeworktextbox = new System.Windows.Controls.TextBox[10];
        
        private void textbox_homeworkcount_TextChanged(object sender, TextChangedEventArgs e)
        {
            

            int lastps = int.Parse(textbox_midtermcount.Text);
            int last = lastps * 30;

            System.Windows.Controls.Label[] homeworklabel = new System.Windows.Controls.Label[int.Parse(textbox_homeworkcount.Text)];
            for (int i = 0; i < int.Parse(textbox_homeworkcount.Text); i++)
            {
                homeworklabel[i] = new System.Windows.Controls.Label();
                homeworklabel[i].Name = "qCountmt" + (i + 1);
                homeworklabel[i].HorizontalAlignment = HorizontalAlignment.Left;
                homeworklabel[i].VerticalAlignment = VerticalAlignment.Top;
                homeworklabel[i].Width = 170;
                homeworklabel[i].Height = 30;
                homeworklabel[i].Opacity = 0.8;
                homeworklabel[i].Content = "Question Count Homework " + (i + 1) + " :";
                homeworklabel[i].Margin = new Thickness(272, (i * 30) + last + 29, 0, 0);
                homeworklabel[i].Foreground = Brushes.White;
                homeworklabel[i].Visibility = Visibility.Visible;
                grid_generate_excel.Children.Add(homeworklabel[i]);
            }

            for (int i = 0; i < int.Parse(textbox_homeworkcount.Text); i++)
            {
                homeworktextbox[i] = new System.Windows.Controls.TextBox();
                homeworktextbox[i].Name = "qTextboxhomework" + (i + 1);
                homeworktextbox[i].HorizontalAlignment = HorizontalAlignment.Left;
                homeworktextbox[i].VerticalAlignment = VerticalAlignment.Top;
                homeworktextbox[i].Width = 70;
                homeworktextbox[i].Height = 15;
                homeworktextbox[i].Margin = new Thickness(440, (i * 30) + last + 35, 0, 0);
                grid_generate_excel.Children.Add(homeworktextbox[i]);
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
        private void checkbox_havefinal_Checked(object sender, RoutedEventArgs e)
        {
            int lastps = int.Parse(textbox_homeworkcount.Text);
            int last = lastps * 30 + 30;

            System.Windows.Controls.Label final_label = new System.Windows.Controls.Label();


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

        private void button_generate_excel_btnr_Click(object sender, RoutedEventArgs e)
        {
            int[] midtermqcount = new int[10];
            for (int i = 0; i < 10; i++)
            {
                if (midtermtextbox[i] == null || int.Parse(midtermtextbox[i].Text) == 0)
                {
                    break;
                }
                else
                {
                    midtermqcount[i] = int.Parse(midtermtextbox[i].Text);
                }
            }
            int[] homeworkqcount = new int[10];
            for (int i = 0; i < 10; i++)
            {
                if (homeworktextbox[i] == null || int.Parse(homeworktextbox[i].Text) == 0)
                {
                    break;
                }
                else
                {
                    homeworkqcount[i] = int.Parse(homeworktextbox[i].Text);
                }
            }

            //Final booleanı girilmeyince hata veriyor
            //Bir değeri diğerinden önnce giremiyorum, örneüin homework sayısı vermeden final durumu işaretlemeyi deneyince hata veriyor.
            //Herhangi bir değeri girip silince hata veriyor
            UserPanel userPanel = new UserPanel();
            userPanel.GenerateExcel(int.Parse(textbox_studentcount.Text), int.Parse(textbox_midtermcount.Text), int.Parse(textbox_homeworkcount.Text), int.Parse(textbox_labcount.Text),
                int.Parse(textbox_quizcount.Text), int.Parse(textbox_projectcount.Text), int.Parse(textbox_derscikticount.Text), checkbox_iscatalog.IsChecked ?? false, checkbox_havefinal.IsChecked ?? false, midtermqcount, homeworkqcount, int.Parse(finaltextbox.Text));
        }

        private void button_createreport_Click(object sender, RoutedEventArgs e)
        {
            
        }

        private void button_downloadexcel_Click(object sender, RoutedEventArgs e)
        {
            GoogleDrive.GetFile();
        }

        private void button_selectexcelfile_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            openFileDialog.ShowDialog();
            string excelfileloc = openFileDialog.FileName;

            List<string> sheetNames = new List<string>();
            List<string> columnNames = new List<string>();


            int midtermsheetcount = 0;
            int []midtermqcount = new int[10];
            bool finalsheet = false;
            int i = 0;
            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelfileloc + "; Extended Properties='Excel 12.0 Xml;HDR=YES'";

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();

                // Get list of sheet names
                System.Data.DataTable schemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                
                foreach (DataRow row in schemaTable.Rows)
                {
                    sheetNames.Add(row["TABLE_NAME"].ToString().Replace("$", ""));
                }

                // Close connection
                connection.Close();
            }
            foreach (string sheetName in sheetNames)
            {
                if (sheetName.Contains("Midterm-"))
                {
                    
                    midtermsheetcount++;

                    using (OleDbConnection connection = new OleDbConnection(connectionString))
                    {
                        connection.Open();

                        System.Data.DataTable dataTable = new System.Data.DataTable();
                        OleDbDataAdapter dataAdapter = new OleDbDataAdapter("SELECT * FROM [" + sheetName + "]", connection);
                        dataAdapter.Fill(dataTable);

                        foreach (DataRow row in dataTable.Rows)
                        {
                            columnNames.Add(row["COLUMN_NAME"].ToString());
                        }


                        connection.Close();
                    }

                    foreach (string columnName in columnNames)
                    {
                        midtermqcount[i]++;
                    }

                }




                if (sheetName.Contains("Final"))
                {
                    finalsheet = true;
                }


                i++;
            }


            MessageBox.Show("a");

        }
    } 
}
