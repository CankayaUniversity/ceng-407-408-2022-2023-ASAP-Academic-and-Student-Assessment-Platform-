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
using ExcelDataReader;
using System.Data;

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

            System.Windows.Controls.TextBox[] midtermtextbox = new System.Windows.Controls.TextBox[int.Parse(textbox_midtermcount.Text)];
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

            System.Windows.Controls.TextBox[] homeworktextbox = new System.Windows.Controls.TextBox[int.Parse(textbox_homeworkcount.Text)];
            for (int i = 0; i < int.Parse(textbox_homeworkcount.Text); i++)
            {
                homeworktextbox[i] = new System.Windows.Controls.TextBox();
                homeworktextbox[i].Name = "qTextboxmidterm" + (i + 1);
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
            //GoogleDrive.UploadFile();
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
        //
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
    } 
}
